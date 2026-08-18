from __future__ import annotations

import csv
import json
import re
import shutil
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent.parent
BASE_CSV = ROOT / "output" / "final_ready" / "doors_catalog.csv"
TARGETS_CSV = ROOT / "data" / "target_lines.csv"
OUT_DIR = ROOT / "output" / "final_ready_v2"
CHECK_DIR = OUT_DIR / "check_by_producer"
OLYMP_CHECK_DIR = ROOT / "output" / "olymp_model_rows_check"

CATALOG_FIELDS = [
    "door_type",
    "producer",
    "target_line",
    "model_name",
    "description",
    "material",
    "coating",
    "image_url",
    "source_url",
    "status",
    "comment",
]

WP_FIELDS = [
    "post_title",
    "post_content",
    "producer",
    "line",
    "material",
    "coating",
    "image_url",
    "source_url",
    "door_type",
    "status",
]

QA_FIELDS = [
    "producer",
    "target_line",
    "expected_material",
    "rows_count",
    "unique_models_count",
    "with_image_count",
    "missing_image_count",
    "needs_review_count",
    "bad_image_count",
    "status",
    "comment",
]

NOT_FOUND_FIELDS = [
    "door_type",
    "manufacturer",
    "target_line",
    "source_site",
    "status",
    "comment",
]

REPARSE_FIELDS = [
    "producer",
    "target_line",
    "model_name",
    "source_url",
    "old_image_url",
    "new_image_url",
    "action",
    "reason",
    "http_status",
    "content_type",
    "width",
    "height",
    "status_after",
    "comment_after",
]

REBUILD_FIELDS = [
    "producer",
    "old_rows",
    "new_rows",
    "old_variant_rows",
    "new_variant_rows",
    "real_model_rows",
    "generated_model_rows",
    "rows_with_image",
    "rows_without_image",
    "repeated_image_gt_10_count",
    "needs_review_count",
    "comment",
]

AUDIT_FIELDS = [
    "producer",
    "rows",
    "rows_with_image",
    "rows_without_image",
    "needs_review",
    "variant_rows",
    "generated_model_rows",
    "repeated_image_gt_10",
    "comment",
]

EXTRACT_FIELDS = [
    "target_line",
    "source_url",
    "card_text",
    "img_alt",
    "img_title",
    "img_filename",
    "image_url",
    "dom_block",
    "is_usable_model",
    "reject_reason",
]


def norm(text: str) -> str:
    text = (text or "").lower().replace("ё", "е")
    text = re.sub(r"[^a-zа-я0-9]+", " ", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip()


def norm_model(text: str) -> str:
    text = norm(text)
    text = re.sub(r"\b(дверь|model|модель|variant|вариант)\b", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def append_comment(comment: str, addon: str) -> str:
    comment = (comment or "").strip()
    addon = (addon or "").strip()
    if not addon:
        return comment
    if not comment:
        return addon
    if addon in comment:
        return comment
    return f"{comment}; {addon}"


def read_csv(path: Path) -> list[dict]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return [{k: (v or "").strip() for k, v in row.items()} for row in csv.DictReader(f, delimiter=";")]


def write_csv(path: Path, rows: list[dict], fields: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields, delimiter=";")
        w.writeheader()
        for row in rows:
            w.writerow({k: row.get(k, "") for k in fields})


def write_xlsx(path: Path, rows: list[dict], fields: list[str], title: str = "parsed_catalog") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = title[:31]
    ws.append(fields)
    for row in rows:
        ws.append([row.get(k, "") for k in fields])
    widths = [14, 22, 28, 40, 70, 18, 16, 60, 60, 14, 42][: len(fields)]
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"
    wb.save(path)


def filename_from_url(url: str) -> str:
    path = urlparse(url).path
    return Path(path).name


def clean_caption(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip()


def is_noise_text(text: str) -> bool:
    t = norm(text)
    if not t:
        return True
    bad_exact = {
        "главная",
        "каталог",
        "где купить",
        "купить",
        "контакты",
        "доставка",
        "оплата",
        "в корзину",
        "подробнее",
    }
    if t in bad_exact:
        return True
    bad_parts = (
        "коммерческого предложения",
        "согласие на обработку",
        "скачайте наши каталоги",
        "политик",
    )
    return any(x in t for x in bad_parts)


def bad_image_url(url: str) -> bool:
    low = (url or "").lower()
    if not low:
        return True
    bad = (
        "logo",
        "icon",
        "badge",
        "sticker",
        "label",
        "sale",
        "promo",
        "banner",
        "search",
        "arrow",
        "placeholder",
        "no_image",
        "no-image",
        "pixel",
        "open.png",
        "catalog.svg",
        "sprite",
        "loader",
        "favicon",
        "/resize/20x/",
        "resize/20x",
        "/cover/240x180/",
        "preview",
        "close.png",
        "tl.png",
        "tg.png",
        "whatsapp",
        "telegram",
    )
    return any(x in low for x in bad)


def meaningful_from_filename(filename: str) -> str:
    stem = Path(filename).stem
    stem = re.sub(r"[_\-]+", " ", stem)
    stem = clean_caption(stem)
    low = norm(stem)
    if not low:
        return ""
    if low in {"photo", "image", "img", "big", "small", "close"}:
        return ""
    if re.fullmatch(r"[\d\s]+", low):
        return ""
    # keep short code names like direkt 1, loft 2, venir 5
    if len(low) < 3:
        return ""
    return stem


def extract_cards_with_playwright(url: str) -> list[dict]:
    script = r"""
    () => {
      const out = [];
      const imgs = Array.from(document.querySelectorAll('img'));
      for (const img of imgs) {
        const src = img.currentSrc || img.getAttribute('data-original') || img.getAttribute('data-src') || img.getAttribute('src') || '';
        if (!src) continue;
        let node = img;
        let domBlock = '';
        let nearText = '';
        for (let i = 0; i < 7 && node; i++) {
          const cls = (node.className || '').toString().trim();
          if (!domBlock && cls) domBlock = cls.split(' ').slice(0, 4).join(' ');
          if (!nearText) {
            const titleNode = node.querySelector ? node.querySelector('.t-name, .t-title, .t-descr, .t-card__title, .t-store__card__title, .tn-atom, [field*="title"], [field*="name"]') : null;
            if (titleNode) {
              const t = (titleNode.textContent || '').trim();
              if (t && t.length <= 120) nearText = t;
            }
          }
          node = node.parentElement;
        }
        out.push({
          image_url: src,
          img_alt: img.getAttribute('alt') || '',
          img_title: img.getAttribute('title') || '',
          card_text: nearText,
          dom_block: domBlock,
          width: img.naturalWidth || 0,
          height: img.naturalHeight || 0
        });
      }
      return out;
    }
    """
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(url, wait_until="domcontentloaded", timeout=90000)
        page.wait_for_timeout(2500)
        rows = page.evaluate(script)
        browser.close()
    return rows


def build_olymp_rows(old_rows: list[dict]) -> tuple[list[dict], list[dict], list[dict], dict]:
    by_line_url: dict[tuple[str, str], list[dict]] = {}
    for row in old_rows:
        by_line_url.setdefault((row.get("target_line", ""), row.get("source_url", "")), []).append(row)

    extracted_rows: list[dict] = []
    unmatched_rows: list[dict] = []
    new_rows: list[dict] = []

    cache: dict[str, list[dict]] = {}
    old_variant = sum(
        1
        for r in old_rows
        if re.search(r"\bvariant\s*\d+\b|\bвариант\s*\d+\b|\bмодель\s*\d+\b", r.get("model_name", ""), flags=re.I)
    )

    for (target_line, source_url), rows in by_line_url.items():
        defaults = rows[0]
        if source_url not in cache:
            cache[source_url] = extract_cards_with_playwright(source_url)
        raw_cards = cache[source_url]

        usable_cards = []
        for card in raw_cards:
            image_url = (card.get("image_url") or "").strip()
            alt = clean_caption(card.get("img_alt", ""))
            title = clean_caption(card.get("img_title", ""))
            card_text = clean_caption(card.get("card_text", ""))
            dom_block = clean_caption(card.get("dom_block", ""))
            w = int(card.get("width") or 0)
            h = int(card.get("height") or 0)

            reject_reason = ""
            if bad_image_url(image_url):
                reject_reason = "bad_image_pattern"
            elif w < 220 or h < 220:
                reject_reason = "too_small_image"

            model_name = ""
            if not reject_reason and card_text and not is_noise_text(card_text):
                model_name = card_text
                source_reason = "rebuilt_from_visible_card_caption"
            elif not reject_reason and alt and not is_noise_text(alt):
                model_name = alt
                source_reason = "rebuilt_from_alt_title_filename"
            elif not reject_reason and title and not is_noise_text(title):
                model_name = title
                source_reason = "rebuilt_from_alt_title_filename"
            elif not reject_reason:
                fn = meaningful_from_filename(filename_from_url(image_url))
                if fn and not is_noise_text(fn):
                    model_name = fn
                    source_reason = "rebuilt_from_alt_title_filename"
                else:
                    source_reason = "no_visible_model_caption"

            is_usable = bool(model_name) and not reject_reason
            if not is_usable and not reject_reason:
                reject_reason = "no_visible_model_caption"

            extracted_rows.append(
                {
                    "target_line": target_line,
                    "source_url": source_url,
                    "card_text": card_text,
                    "img_alt": alt,
                    "img_title": title,
                    "img_filename": filename_from_url(image_url),
                    "image_url": image_url,
                    "dom_block": dom_block,
                    "is_usable_model": "1" if is_usable else "0",
                    "reject_reason": reject_reason,
                }
            )

            if is_usable:
                usable_cards.append(
                    {
                        "model_name": clean_caption(model_name),
                        "image_url": image_url,
                        "comment": source_reason,
                    }
                )
            else:
                unmatched_rows.append(
                    {
                        "target_line": target_line,
                        "source_url": source_url,
                        "card_text": card_text,
                        "img_alt": alt,
                        "img_title": title,
                        "img_filename": filename_from_url(image_url),
                        "image_url": image_url,
                        "dom_block": dom_block,
                        "is_usable_model": "0",
                        "reject_reason": reject_reason or "rejected",
                    }
                )

        # dedupe usable
        uniq = {}
        for c in usable_cards:
            key = (norm(c["model_name"]), c["image_url"])
            if key not in uniq:
                uniq[key] = c
        usable_cards = list(uniq.values())

        if usable_cards:
            for c in usable_cards:
                new_rows.append(
                    {
                        "door_type": defaults.get("door_type", "межкомнатная"),
                        "producer": "Олимп",
                        "target_line": target_line,
                        "model_name": c["model_name"],
                        "description": defaults.get("description", ""),
                        "material": defaults.get("material", ""),
                        "coating": defaults.get("coating", ""),
                        "image_url": c["image_url"],
                        "source_url": source_url,
                        "status": "found",
                        "comment": c["comment"],
                    }
                )
        else:
            # Only one collection row if no visible model captions
            fallback = ""
            for card in raw_cards:
                img = (card.get("image_url") or "").strip()
                if not bad_image_url(img):
                    fallback = img
                    break
            new_rows.append(
                {
                    "door_type": defaults.get("door_type", "межкомнатная"),
                    "producer": "Олимп",
                    "target_line": target_line,
                    "model_name": target_line,
                    "description": defaults.get("description", ""),
                    "material": defaults.get("material", ""),
                    "coating": defaults.get("coating", ""),
                    "image_url": fallback,
                    "source_url": source_url,
                    "status": "needs_review",
                    "comment": "collection_image_only_no_visible_model_caption",
                }
            )

    img_counts = Counter(r.get("image_url", "") for r in new_rows if r.get("image_url", ""))
    repeated_gt10 = 0
    for r in new_rows:
        img = r.get("image_url", "")
        if img and img_counts.get(img, 0) > 10:
            repeated_gt10 += 1
            r["status"] = "needs_review"
            r["comment"] = append_comment(r.get("comment", ""), "collection_image_used")

    stats = {
        "old_rows": len(old_rows),
        "new_rows": len(new_rows),
        "old_variant_rows": old_variant,
        "new_variant_rows": sum(
            1
            for r in new_rows
            if re.search(r"\bvariant\s*\d+\b|\bвариант\s*\d+\b|\bмодель\s*\d+\b", r.get("model_name", ""), flags=re.I)
        ),
        "real_model_rows": sum(
            1
            for r in new_rows
            if "rebuilt_from_visible_card_caption" in (r.get("comment") or "") or "rebuilt_from_alt_title_filename" in (r.get("comment") or "")
        ),
        "generated_model_rows": sum(
            1 for r in new_rows if "collection_image_only_no_visible_model_caption" in (r.get("comment") or "")
        ),
        "rows_with_image": sum(1 for r in new_rows if r.get("image_url", "")),
        "rows_without_image": sum(1 for r in new_rows if not r.get("image_url", "")),
        "repeated_image_gt_10_count": repeated_gt10,
        "needs_review_count": sum(1 for r in new_rows if r.get("status") == "needs_review"),
    }
    return new_rows, extracted_rows, unmatched_rows, stats


def recalc_qa(rows: list[dict], target_rows: list[dict]) -> tuple[list[dict], list[dict], dict]:
    grouped: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for row in rows:
        grouped[(norm(row.get("producer", "")), norm(row.get("target_line", "")))].append(row)

    qa_rows: list[dict] = []
    not_found_rows: list[dict] = []
    lines_total = 0
    lines_real = 0
    lines_placeholder = 0
    lines_not_found = 0
    for t in target_rows:
        producer = (t.get("manufacturer") or "").strip()
        line = (t.get("target_line") or "").strip()
        if not producer or not line:
            continue
        lines_total += 1
        key = (norm(producer), norm(line))
        sub = grouped.get(key, [])
        if not sub and norm(producer) == norm("АСД") and norm(line) == norm("Все модели"):
            sub = grouped.get((norm("АСД"), norm("Уточнить коллекцию")), [])

        with_image = sum(1 for r in sub if r.get("image_url", "").strip())
        missing = len(sub) - with_image
        needs_review = sum(1 for r in sub if (r.get("status") or "").strip() == "needs_review")
        bad_img = sum(1 for r in sub if bad_image_url(r.get("image_url", "")))
        unique_models = len({norm_model(r.get("model_name", "")) for r in sub if r.get("model_name", "").strip()})

        placeholders = [r for r in sub if "collection_not_found_on_site" in (r.get("comment") or "").lower()]
        real_rows = [r for r in sub if r not in placeholders]
        if not sub:
            status = "not_found"
            comment = "checked_catalog_no_rows_for_target_line"
            lines_not_found += 1
            not_found_rows.append(
                {
                    "door_type": t.get("door_type", ""),
                    "manufacturer": producer,
                    "target_line": line,
                    "source_site": t.get("source_site", ""),
                    "status": "not_found",
                    "comment": comment,
                }
            )
        elif real_rows:
            lines_real += 1
            status = "needs_review" if (missing > 0 or needs_review > 0 or bad_img > 0) else "found"
            comment = f"with_image={with_image}; missing_image={missing}; needs_review={needs_review}; bad_image={bad_img}"
        else:
            lines_placeholder += 1
            status = "needs_review"
            comment = "collection_not_found_on_site"

        qa_rows.append(
            {
                "producer": producer,
                "target_line": line,
                "expected_material": (t.get("covering_from_source") or t.get("material_manual") or "").strip(),
                "rows_count": len(sub),
                "unique_models_count": unique_models,
                "with_image_count": with_image,
                "missing_image_count": missing,
                "needs_review_count": needs_review,
                "bad_image_count": bad_img,
                "status": status,
                "comment": comment,
            }
        )
    coverage = {
        "target_lines_total": lines_total,
        "target_lines_with_real_models": lines_real,
        "target_lines_as_placeholder": lines_placeholder,
        "target_lines_not_found": lines_not_found,
    }
    return qa_rows, not_found_rows, coverage


def main() -> None:
    if OUT_DIR.exists():
        shutil.rmtree(OUT_DIR)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    CHECK_DIR.mkdir(parents=True, exist_ok=True)
    OLYMP_CHECK_DIR.mkdir(parents=True, exist_ok=True)

    base_rows = read_csv(BASE_CSV)
    targets = read_csv(TARGETS_CSV)

    olymp_old = [r for r in base_rows if r.get("producer") == "Олимп"]
    non_olymp = [r for r in base_rows if r.get("producer") != "Олимп"]

    olymp_new, extracted, unmatched, olymp_stats = build_olymp_rows(olymp_old)
    write_csv(OLYMP_CHECK_DIR / "olymp_extracted_cards.csv", extracted, EXTRACT_FIELDS)
    write_csv(OLYMP_CHECK_DIR / "olymp_unmatched_images.csv", unmatched, EXTRACT_FIELDS)

    final_rows = non_olymp + olymp_new
    rows_sorted = sorted(final_rows, key=lambda r: (norm(r.get("producer", "")), norm(r.get("target_line", "")), norm(r.get("model_name", ""))))

    write_csv(OUT_DIR / "doors_catalog.csv", rows_sorted, CATALOG_FIELDS)
    write_xlsx(OUT_DIR / "doors_catalog.xlsx", rows_sorted, CATALOG_FIELDS)
    wp_rows = [
        {
            "post_title": r.get("model_name", ""),
            "post_content": r.get("description", ""),
            "producer": r.get("producer", ""),
            "line": r.get("target_line", ""),
            "material": r.get("material", ""),
            "coating": r.get("coating", ""),
            "image_url": r.get("image_url", ""),
            "source_url": r.get("source_url", ""),
            "door_type": r.get("door_type", ""),
            "status": r.get("status", ""),
        }
        for r in rows_sorted
    ]
    write_csv(OUT_DIR / "wordpress_import.csv", wp_rows, WP_FIELDS)

    qa_rows, not_found_rows, coverage = recalc_qa(rows_sorted, targets)
    write_csv(OUT_DIR / "qa_report.csv", qa_rows, QA_FIELDS)
    write_csv(OUT_DIR / "not_found.csv", not_found_rows, NOT_FOUND_FIELDS)

    # image_reparse for Olymp
    image_reparse_rows: list[dict] = []
    old_map = {(r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", "")): r for r in olymp_old}
    new_map = {(r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", "")): r for r in olymp_new}
    keys = set(old_map) | set(new_map)
    for key in sorted(keys):
        o = old_map.get(key)
        n = new_map.get(key)
        if o and not n:
            action = "removed_old_row"
            reason = "old_generated_model_removed"
            row = {
                "producer": "Олимп",
                "target_line": o.get("target_line", ""),
                "model_name": o.get("model_name", ""),
                "source_url": o.get("source_url", ""),
                "old_image_url": o.get("image_url", ""),
                "new_image_url": "",
                "action": action,
                "reason": reason,
                "http_status": "",
                "content_type": "",
                "width": "",
                "height": "",
                "status_after": "",
                "comment_after": "",
            }
        elif n and not o:
            reason = "rebuilt_from_visible_card_caption" if "rebuilt_from_visible_card_caption" in (n.get("comment") or "") else "rebuilt_from_alt_title_filename"
            if "collection_image_only_no_visible_model_caption" in (n.get("comment") or ""):
                reason = "collection_image_only_no_visible_model_caption"
            row = {
                "producer": "Олимп",
                "target_line": n.get("target_line", ""),
                "model_name": n.get("model_name", ""),
                "source_url": n.get("source_url", ""),
                "old_image_url": "",
                "new_image_url": n.get("image_url", ""),
                "action": "added_rebuilt_row",
                "reason": reason,
                "http_status": "",
                "content_type": "",
                "width": "",
                "height": "",
                "status_after": n.get("status", ""),
                "comment_after": n.get("comment", ""),
            }
        else:
            if (
                o.get("image_url", "") == n.get("image_url", "")
                and o.get("status", "") == n.get("status", "")
                and o.get("comment", "") == n.get("comment", "")
            ):
                continue
            row = {
                "producer": "Олимп",
                "target_line": n.get("target_line", ""),
                "model_name": n.get("model_name", ""),
                "source_url": n.get("source_url", ""),
                "old_image_url": o.get("image_url", ""),
                "new_image_url": n.get("image_url", ""),
                "action": "updated_row",
                "reason": "image_or_status_updated",
                "http_status": "",
                "content_type": "",
                "width": "",
                "height": "",
                "status_after": n.get("status", ""),
                "comment_after": n.get("comment", ""),
            }
        image_reparse_rows.append(row)
    write_csv(OUT_DIR / "image_reparse_report.csv", image_reparse_rows, REPARSE_FIELDS)

    # rebuild report
    rebuild_rows = [
        {
            "producer": "Олимп",
            **olymp_stats,
            "comment": "rebuilt_olymp_from_visible_cards_and_alt_title_filename",
        }
    ]
    write_csv(OUT_DIR / "rebuild_report.csv", rebuild_rows, REBUILD_FIELDS)

    # check_by_producer all 10
    producer_file = {
        "Скамбио Порте": "scambio",
        "Олимп": "olymp",
        "Хаусдорс": "hausdoors",
        "Лорд": "lord",
        "Аэлита": "aelita",
        "Энтро": "entro",
        "Комфорт": "komfort",
        "Двери регионов": "dveriregionov",
        "Термодур": "termodur",
        "АСД": "asd",
    }
    producers = sorted({r.get("producer", "") for r in rows_sorted if r.get("producer", "")}, key=lambda x: norm(x))
    audit_rows: list[dict] = []
    for p in producers:
        sub = [r for r in rows_sorted if r.get("producer") == p]
        stem = producer_file.get(p, norm(p).replace(" ", "_"))
        write_csv(CHECK_DIR / f"{stem}.csv", sub, CATALOG_FIELDS)
        write_xlsx(CHECK_DIR / f"{stem}.xlsx", sub, CATALOG_FIELDS, title=stem)
        img_counts = Counter(r.get("image_url", "") for r in sub if r.get("image_url", ""))
        repeated_gt10 = sum(1 for r in sub if r.get("image_url", "") and img_counts.get(r.get("image_url", ""), 0) > 10)
        variant_rows = sum(
            1 for r in sub if re.search(r"\bvariant\s*\d+\b|\bвариант\s*\d+\b|\bмодель\s*\d+\b", r.get("model_name", ""), flags=re.I)
        )
        generated_rows = sum(
            1
            for r in sub
            if "collection_image_only_no_visible_model_caption" in (r.get("comment") or "")
            or "generated_model_name_no_visible_caption" in (r.get("comment") or "")
        )
        audit_rows.append(
            {
                "producer": p,
                "rows": len(sub),
                "rows_with_image": sum(1 for r in sub if r.get("image_url", "")),
                "rows_without_image": sum(1 for r in sub if not r.get("image_url", "")),
                "needs_review": sum(1 for r in sub if r.get("status") == "needs_review"),
                "variant_rows": variant_rows,
                "generated_model_rows": generated_rows,
                "repeated_image_gt_10": repeated_gt10,
                "comment": "olymp_rebuilt" if p == "Олимп" else "unchanged",
            }
        )
    write_csv(CHECK_DIR / "final_problem_sites_audit.csv", audit_rows, AUDIT_FIELDS)

    # run report
    base_by_prod = defaultdict(list)
    new_by_prod = defaultdict(list)
    for r in base_rows:
        base_by_prod[r.get("producer", "")].append(r)
    for r in rows_sorted:
        new_by_prod[r.get("producer", "")].append(r)
    unchanged = {}
    for p in ("Энтро", "Хаусдорс", "АСД", "Скамбио Порте", "Лорд", "Аэлита", "Комфорт", "Двери регионов", "Термодур"):
        b = sorted(base_by_prod.get(p, []), key=lambda r: (r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", ""), r.get("image_url", "")))
        n = sorted(new_by_prod.get(p, []), key=lambda r: (r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", ""), r.get("image_url", "")))
        unchanged[p] = {"rows_before": len(b), "rows_after": len(n), "exact_match": b == n}

    run_report = {
        "generated_at": datetime.now().isoformat(),
        "mode": "final_ready_v2_olymp_only",
        "source_catalog": str(BASE_CSV),
        "old_total_rows": len(base_rows),
        "total_rows": len(rows_sorted),
        "producers_count": len(producers),
        **coverage,
        "rows_with_image": sum(1 for r in rows_sorted if r.get("image_url", "")),
        "rows_without_image": sum(1 for r in rows_sorted if not r.get("image_url", "")),
        "needs_review_count": sum(1 for r in rows_sorted if r.get("status") == "needs_review"),
        "rows_by_producer": dict(sorted(Counter(r.get("producer", "") for r in rows_sorted if r.get("producer", "")).items())),
        "olymp_rebuild_stats": olymp_stats,
        "olymp_unmatched_images_count": len(unmatched),
        "unchanged_producers_check": unchanged,
    }
    (OUT_DIR / "run_report.json").write_text(json.dumps(run_report, ensure_ascii=False, indent=2), encoding="utf-8")

    print("Built output/final_ready_v2")
    print(json.dumps({"old_olymp_rows": len(olymp_old), "new_olymp_rows": len(olymp_new), **olymp_stats, "unmatched": len(unmatched)}, ensure_ascii=False))


if __name__ == "__main__":
    main()

