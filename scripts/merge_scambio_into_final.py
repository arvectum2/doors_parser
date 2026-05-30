from __future__ import annotations

import argparse
import csv
import json
import re
import shutil
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


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

IMAGE_AUDIT_FIELDS = [
    "producer",
    "target_line",
    "model_name",
    "old_image_url",
    "new_image_url",
    "source_url",
    "image_status_before",
    "image_status_after",
    "action",
    "comment",
]

DUP_FIELDS = [
    "producer",
    "target_line",
    "normalized_model_name",
    "duplicates_count_before",
    "kept_model_name",
    "kept_image_url",
    "removed_image_urls",
    "reason",
]

AUDIT_SUMMARY_FIELDS = [
    "producer",
    "rows_count",
    "target_lines_count",
    "missing_image_count",
    "needs_review_count",
    "material_filled_count",
    "suspicious_image_count",
    "suspicious_model_name_count",
    "duplicate_model_name_count",
    "comment",
]

PRODUCER_TO_FILE = {
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

FORBIDDEN_MODEL_PREFIXES = (
    "каталог",
    "поиск",
    "меню",
    "корзина",
    "избранное",
    "сравнение",
    "комплект",
    "коробка",
    "наличник",
    "добор",
    "погонаж",
    "телескоп",
    "притвор",
    "капитель",
    "карниз",
    "розетка",
    "плинтус",
    "подпятник",
    "профиль",
    "разрез",
    "сечение",
    "образец",
    "цвет",
    "покрытие",
)

CRITICAL_BAD_IMAGE_HINTS = (
    "logo",
    "icon",
    "catalog.svg",
    "search_loop",
    "colors_images",
    "open.png",
    "no_image",
    "no-image",
    "noimage",
    "placeholder",
    "blank",
    "sprite",
    "loader",
    "pixel",
    "pixel.gif",
    "favicon",
    "thumb",
    "thumbnail",
    "/96x",
    "/100x",
    "/150x",
    "/200x",
)

SUSPICIOUS_IMAGE_HINTS = (
    "logo",
    "icon",
    "badge",
    "label",
    "sale",
    "new",
    "hit",
    "cart",
    "search",
    "arrow",
    "marker",
    "promo",
    "sticker",
    "catalog.svg",
    "search_loop",
    "colors_images",
    "open.png",
    "no_image",
    "no-image",
    "noimage",
    "placeholder",
    "blank",
    "sprite",
    "loader",
    "pixel",
    "pixel.gif",
    "favicon",
    "thumb",
    "thumbnail",
    "small",
    "mini",
    "preview",
    "resize",
    "96x",
    "100x",
    "150x",
    "200x",
)

SUSPICIOUS_MODEL_TOKENS = FORBIDDEN_MODEL_PREFIXES

PRICE_PATTERNS = (
    r"₽",
    r"\bprice\b",
    r"цена\s+указана",
    r"\bруб(лей|\.|ля)?\b",
    r"от\s*\d[\d\s]{2,}",
)

SCAMBIO_LINES_REQUIRED = [
    "Смальта Деко",
    "Смальта Белла",
    "Смальта Галант",
    "Смальта Лайн",
    "Смальта Нова",
    "Смальта Риф",
    "Смальта Шарм",
    "ПС",
    "ПСЛ",
    "ПСК",
    "ПСЦ Модерн",
    "ПСЦ Классика",
    "ПСУ",
    "ПХ Алл кромка Черный",
    "ПХ Алл кромка Хром",
    "ПСТ",
    "ПСЕ",
    "ПДЕ",
    "П АБС кромка Черный",
    "П",
    "Инвизибл",
]


@dataclass
class MergeStats:
    bad_images_removed: int = 0
    model_rows_removed: int = 0
    model_rows_cleaned: int = 0
    duplicates_removed: int = 0


def normalize(text: str) -> str:
    text = (text or "").lower().replace("ё", "е")
    text = re.sub(r"[^a-zа-я0-9]+", " ", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip()


def normalize_model_key(text: str) -> str:
    t = normalize(text)
    t = re.sub(r"\b(дверь|модель)\b", " ", t)
    return re.sub(r"\s+", " ", t).strip()


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

    widths = [14, 22, 28, 40, 70, 18, 16, 60, 60, 14, 40][: len(fields)]
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


def bad_image_critical(url: str) -> bool:
    low = (url or "").lower()
    if not low:
        return False
    return any(h in low for h in CRITICAL_BAD_IMAGE_HINTS)


def suspicious_image(url: str) -> bool:
    low = (url or "").lower()
    if not low:
        return True
    return any(h in low for h in SUSPICIOUS_IMAGE_HINTS)


def suspicious_model_name(name: str) -> bool:
    low = normalize(name)
    if not low:
        return True
    return any(tok in low for tok in SUSPICIOUS_MODEL_TOKENS)


def clean_model_name(name: str) -> str:
    t = name or ""
    t = re.sub(r"\s*[–—-]?\s*распродажа\s*[.,;:!?]*", " ", t, flags=re.IGNORECASE)
    t = t.replace("×", "x").replace("Х", "x").replace("х", "x")
    t = re.sub(r"\bразмер\s*[:\-]?\s*\d[\d.]*\s*x\s*\d[\d.]*(\s*x\s*\d[\d.]*)?\b", " ", t, flags=re.I)
    t = re.sub(r"\b\d[\d.]*\s*x\s*\d[\d.]*(\s*x\s*\d[\d.]*)?\b", " ", t, flags=re.I)
    t = re.sub(r"\.{2,}", " ", t)
    t = re.sub(r"\s+", " ", t).strip(" \t\r\n.,;:-")
    return t


def has_price_content(text: str) -> bool:
    low = (text or "").lower()
    return any(re.search(p, low) for p in PRICE_PATTERNS)


def clean_description(text: str) -> str:
    if not text:
        return ""
    if not has_price_content(text):
        return text
    chunks = re.split(r"(?<=[.!?])\s+", text)
    kept = [c for c in chunks if not has_price_content(c)]
    cleaned = " ".join(kept).strip()
    return cleaned or ""


def model_starts_forbidden(name: str) -> bool:
    low = normalize(name)
    if not low:
        return True
    return any(low == p or low.startswith(p + " ") for p in FORBIDDEN_MODEL_PREFIXES)


def row_score(row: dict) -> int:
    score = 0
    if (row.get("status") or "").strip() == "found":
        score += 60
    if (row.get("image_url") or "").strip():
        score += 40
    if "/xml/images/" in (row.get("image_url") or "").lower():
        score += 30
    if (row.get("material") or "").strip():
        score += 10
    if "needs_review" in (row.get("status") or "").lower():
        score -= 5
    return score


def row_is_placeholder(row: dict) -> bool:
    c = (row.get("comment") or "").lower()
    if "collection_not_found_on_site" in c:
        return True
    if (row.get("status") or "") == "needs_review" and not (row.get("image_url") or "").strip():
        if normalize(row.get("model_name", "")) == normalize(row.get("target_line", "")):
            return True
    return False


def ensure_catalog_fields(rows: list[dict]) -> list[dict]:
    out: list[dict] = []
    for r in rows:
        row = {k: (r.get(k, "") or "").strip() for k in CATALOG_FIELDS}
        out.append(row)
    return out


def validate_scambio_rows(scambio_rows: list[dict]) -> dict:
    lines = Counter(r["target_line"] for r in scambio_rows)
    bad_name = [
        r for r in scambio_rows
        if re.search(r"распродажа|размер|\.\.\.", (r.get("model_name") or "").lower())
        or model_starts_forbidden(r.get("model_name", ""))
    ]
    report = {
        "rows_total": len(scambio_rows),
        "line_count": len(lines),
        "lines_present": sorted(lines.keys()),
        "required_lines_missing": [ln for ln in SCAMBIO_LINES_REQUIRED if ln not in lines],
        "rows_with_image": sum(1 for r in scambio_rows if r.get("image_url", "").strip()),
        "rows_without_image": sum(1 for r in scambio_rows if not r.get("image_url", "").strip()),
        "bad_model_rows": len(bad_name),
    }
    return report


def deduplicate_rows(rows: list[dict]) -> tuple[list[dict], list[dict], int]:
    grouped: dict[tuple[str, str, str], list[dict]] = defaultdict(list)
    for r in rows:
        key = (
            normalize(r.get("producer", "")),
            normalize(r.get("target_line", "")),
            normalize_model_key(r.get("model_name", "")),
        )
        grouped[key].append(r)

    selected: list[dict] = []
    dup_rows: list[dict] = []
    removed = 0
    for (_, _, norm_name), group in grouped.items():
        ranked = sorted(group, key=row_score, reverse=True)
        keep = ranked[0]
        selected.append(keep)
        if len(ranked) > 1:
            removed += len(ranked) - 1
            removed_images = [g.get("image_url", "") for g in ranked[1:] if g.get("image_url", "").strip()]
            dup_rows.append(
                {
                    "producer": keep.get("producer", ""),
                    "target_line": keep.get("target_line", ""),
                    "normalized_model_name": norm_name,
                    "duplicates_count_before": len(group),
                    "kept_model_name": keep.get("model_name", ""),
                    "kept_image_url": keep.get("image_url", ""),
                    "removed_image_urls": " | ".join(removed_images),
                    "reason": "duplicate_model_same_line",
                }
            )
    return selected, dup_rows, removed


def build_wordpress_rows(rows: list[dict]) -> list[dict]:
    out = []
    for r in rows:
        out.append(
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
        )
    return out


def build_target_maps(target_rows: list[dict]) -> tuple[dict[tuple[str, str], dict], list[dict]]:
    target_map: dict[tuple[str, str], dict] = {}
    normalized_targets: list[dict] = []
    for t in target_rows:
        producer = (t.get("manufacturer") or "").strip()
        line = (t.get("target_line") or "").strip()
        if not producer or not line:
            continue
        rec = {
            "door_type": (t.get("door_type") or "").strip(),
            "manufacturer": producer,
            "target_line": line,
            "expected_material": (t.get("covering_from_source") or t.get("material_manual") or "").strip(),
            "source_site": (t.get("source_site") or "").strip(),
        }
        target_map[(normalize(producer), normalize(line))] = rec
        normalized_targets.append(rec)
    return target_map, normalized_targets


def build_qa_and_not_found(rows: list[dict], target_rows: list[dict]) -> tuple[list[dict], list[dict], dict]:
    grouped: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for r in rows:
        grouped[(normalize(r.get("producer", "")), normalize(r.get("target_line", "")))].append(r)

    qa_rows: list[dict] = []
    not_found_rows: list[dict] = []

    target_lines_total = 0
    lines_real = 0
    lines_placeholder = 0
    lines_not_found = 0

    for t in target_rows:
        producer = t["manufacturer"]
        line = t["target_line"]
        key = (normalize(producer), normalize(line))
        sub = grouped.get(key, [])
        # Legacy target for ASD uses "Все модели", while final catalog keeps "Уточнить коллекцию".
        if not sub and normalize(producer) == normalize("АСД") and normalize(line) == normalize("Все модели"):
            sub = grouped.get((normalize("АСД"), normalize("Уточнить коллекцию")), [])
        target_lines_total += 1

        with_image = sum(1 for r in sub if r.get("image_url", "").strip())
        missing_image = len(sub) - with_image
        needs_review = sum(1 for r in sub if (r.get("status") or "").strip() == "needs_review")
        bad_images = sum(1 for r in sub if bad_image_critical(r.get("image_url", "")))
        unique_models = len({normalize_model_key(r.get("model_name", "")) for r in sub if r.get("model_name", "").strip()})

        placeholders = [r for r in sub if row_is_placeholder(r)]
        real_rows = [r for r in sub if not row_is_placeholder(r)]

        if not sub:
            status = "not_found"
            comment = "checked_catalog_no_rows_for_target_line"
            lines_not_found += 1
            not_found_rows.append(
                {
                    "door_type": t["door_type"],
                    "manufacturer": producer,
                    "target_line": line,
                    "source_site": t.get("source_site", ""),
                    "status": "not_found",
                    "comment": "checked_catalog_no_rows_for_target_line",
                }
            )
        elif real_rows:
            lines_real += 1
            if needs_review > 0 or missing_image > 0 or bad_images > 0:
                status = "needs_review"
                comment = f"with_image={with_image}; missing_image={missing_image}; needs_review={needs_review}; bad_image={bad_images}"
            else:
                status = "found"
                comment = "ok"
        else:
            lines_placeholder += 1
            status = "needs_review"
            comment = "collection_not_found_on_site"

        qa_rows.append(
            {
                "producer": producer,
                "target_line": line,
                "expected_material": t.get("expected_material", ""),
                "rows_count": len(sub),
                "unique_models_count": unique_models,
                "with_image_count": with_image,
                "missing_image_count": missing_image,
                "needs_review_count": needs_review,
                "bad_image_count": bad_images,
                "status": status,
                "comment": comment,
            }
        )

    coverage = {
        "target_lines_total": target_lines_total,
        "target_lines_with_real_models": lines_real,
        "target_lines_as_placeholder": lines_placeholder,
        "target_lines_not_found": lines_not_found,
    }
    return qa_rows, not_found_rows, coverage


def write_samples_for_visual_check(path: Path, by_producer: dict[str, list[dict]]) -> None:
    wb = Workbook()
    ws0 = wb.active
    wb.remove(ws0)

    for producer, rows in by_producer.items():
        if producer not in PRODUCER_TO_FILE:
            continue
        ws = wb.create_sheet(PRODUCER_TO_FILE[producer][:31])
        fields = ["sample_group"] + CATALOG_FIELDS
        ws.append(fields)

        first20 = rows[:20]
        with_image = [r for r in rows if r.get("image_url", "").strip()][:20]
        needs_review = [r for r in rows if (r.get("status") or "").strip() == "needs_review"][:20]
        suspicious = [r for r in rows if suspicious_image(r.get("image_url", "")) or suspicious_model_name(r.get("model_name", ""))][:20]

        for group, chunk in [
            ("first_20", first20),
            ("with_image_20", with_image),
            ("needs_review_20", needs_review),
            ("suspicious_20", suspicious),
        ]:
            for row in chunk:
                ws.append([group] + [row.get(k, "") for k in CATALOG_FIELDS])

        widths = [18, 14, 22, 28, 40, 70, 18, 16, 60, 60, 14, 40]
        for col, width in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = width
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="4F81BD")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    wb.save(path)


def build_check_by_producer(rows: list[dict], out_dir: Path) -> list[dict]:
    cbp = out_dir / "check_by_producer"
    cbp.mkdir(parents=True, exist_ok=True)

    by_producer: dict[str, list[dict]] = defaultdict(list)
    for row in rows:
        by_producer[row.get("producer", "")].append(row)

    summary: list[dict] = []
    for producer, filename in PRODUCER_TO_FILE.items():
        sub = sorted(
            by_producer.get(producer, []),
            key=lambda r: (normalize(r.get("target_line", "")), normalize(r.get("model_name", ""))),
        )
        write_csv(cbp / f"{filename}.csv", sub, CATALOG_FIELDS)
        write_xlsx(cbp / f"{filename}.xlsx", sub, CATALOG_FIELDS, title=filename)

        missing_image = sum(1 for r in sub if not r.get("image_url", "").strip())
        needs_review = sum(1 for r in sub if (r.get("status") or "").strip() == "needs_review")
        material_filled = sum(1 for r in sub if r.get("material", "").strip())
        suspicious_images = sum(1 for r in sub if suspicious_image(r.get("image_url", "")))
        suspicious_models = sum(1 for r in sub if suspicious_model_name(r.get("model_name", "")))

        dup_counter = Counter(
            (normalize(r.get("target_line", "")), normalize_model_key(r.get("model_name", "")))
            for r in sub
            if r.get("model_name", "").strip()
        )
        dup_count = sum(cnt - 1 for cnt in dup_counter.values() if cnt > 1)

        summary.append(
            {
                "producer": producer,
                "rows_count": len(sub),
                "target_lines_count": len({normalize(r.get("target_line", "")) for r in sub if r.get("target_line", "").strip()}),
                "missing_image_count": missing_image,
                "needs_review_count": needs_review,
                "material_filled_count": material_filled,
                "suspicious_image_count": suspicious_images,
                "suspicious_model_name_count": suspicious_models,
                "duplicate_model_name_count": dup_count,
                "comment": "ok" if sub else "no_rows_in_final_catalog",
            }
        )

    write_csv(cbp / "audit_summary.csv", summary, AUDIT_SUMMARY_FIELDS)
    write_samples_for_visual_check(cbp / "samples_for_visual_check.xlsx", by_producer)
    return summary


def clean_and_filter_rows(rows: list[dict]) -> tuple[list[dict], list[dict], MergeStats]:
    stats = MergeStats()
    cleaned_rows: list[dict] = []
    image_audit_rows: list[dict] = []

    for row in rows:
        old_model = row.get("model_name", "")
        old_desc = row.get("description", "")
        old_image = row.get("image_url", "")
        old_status = row.get("status", "")
        comment = row.get("comment", "")

        model_name = clean_model_name(old_model)
        if model_name != old_model:
            stats.model_rows_cleaned += 1
        if model_starts_forbidden(model_name) or not model_name.strip():
            stats.model_rows_removed += 1
            continue

        description = clean_description(old_desc)
        image_url = old_image.strip()
        new_status = old_status or "needs_review"
        action = "keep"
        audit_comment = ""

        if image_url and bad_image_critical(image_url):
            image_url = ""
            new_status = "needs_review"
            action = "bad_image_removed"
            audit_comment = "critical_bad_image_url_removed"
            stats.bad_images_removed += 1
            if "bad_image_removed" not in comment:
                comment = (comment + "; bad_image_removed").strip("; ").strip()

        cleaned = {
            "door_type": row.get("door_type", ""),
            "producer": row.get("producer", ""),
            "target_line": row.get("target_line", ""),
            "model_name": model_name,
            "description": description,
            "material": row.get("material", ""),
            "coating": row.get("coating", ""),
            "image_url": image_url,
            "source_url": row.get("source_url", ""),
            "status": new_status,
            "comment": comment,
        }
        cleaned_rows.append(cleaned)

        if action != "keep":
            image_audit_rows.append(
                {
                    "producer": cleaned["producer"],
                    "target_line": cleaned["target_line"],
                    "model_name": cleaned["model_name"],
                    "old_image_url": old_image,
                    "new_image_url": image_url,
                    "source_url": cleaned["source_url"],
                    "image_status_before": old_status,
                    "image_status_after": new_status,
                    "action": action,
                    "comment": audit_comment,
                }
            )
    return cleaned_rows, image_audit_rows, stats


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="Merge improved Scambio into final client package without reparsing all sites.")
    ap.add_argument("--base-catalog", type=Path, default=Path("output/doors_catalog.csv"))
    ap.add_argument("--scambio-catalog", type=Path, default=Path("output/scambio_check/doors_catalog.csv"))
    ap.add_argument("--targets", type=Path, default=Path("data/target_lines.csv"))
    ap.add_argument("--out", type=Path, default=Path("output/final_merged"))
    return ap.parse_args()


def main() -> None:
    args = parse_args()
    if args.out.exists():
        shutil.rmtree(args.out)
    args.out.mkdir(parents=True, exist_ok=True)

    base_rows_raw = read_csv(args.base_catalog)
    scambio_rows_raw = read_csv(args.scambio_catalog)
    base_rows = ensure_catalog_fields(base_rows_raw)
    scambio_rows = ensure_catalog_fields(scambio_rows_raw)

    scambio_check = validate_scambio_rows(scambio_rows)

    base_without_scambio = [r for r in base_rows if normalize(r.get("producer", "")) != normalize("Скамбио Порте")]
    merged = base_without_scambio + scambio_rows

    cleaned_rows, image_audit_rows, stats = clean_and_filter_rows(merged)
    dedup_rows, duplicate_rows, dup_removed = deduplicate_rows(cleaned_rows)
    stats.duplicates_removed = dup_removed

    # keep stable order by producer/line/model for easier manual QA
    dedup_rows = sorted(
        dedup_rows,
        key=lambda r: (normalize(r.get("producer", "")), normalize(r.get("target_line", "")), normalize(r.get("model_name", ""))),
    )

    # QA and not_found from target table
    target_rows_raw = read_csv(args.targets)
    _, target_rows = build_target_maps(target_rows_raw)
    qa_rows, not_found_rows, coverage = build_qa_and_not_found(dedup_rows, target_rows)

    # producer-level checks
    producer_summary = build_check_by_producer(dedup_rows, args.out)

    # Write final_merged files
    write_csv(args.out / "doors_catalog.csv", dedup_rows, CATALOG_FIELDS)
    write_xlsx(args.out / "doors_catalog.xlsx", dedup_rows, CATALOG_FIELDS)
    wp_rows = build_wordpress_rows(dedup_rows)
    write_csv(args.out / "wordpress_import.csv", wp_rows, WP_FIELDS)
    write_csv(args.out / "qa_report.csv", qa_rows, QA_FIELDS)
    write_csv(args.out / "not_found.csv", not_found_rows, NOT_FOUND_FIELDS)
    write_csv(args.out / "image_audit.csv", image_audit_rows, IMAGE_AUDIT_FIELDS)
    write_csv(args.out / "duplicates_report.csv", duplicate_rows, DUP_FIELDS)

    producers = {r.get("producer", "") for r in dedup_rows if r.get("producer", "").strip()}
    rows_with_image = sum(1 for r in dedup_rows if r.get("image_url", "").strip())
    rows_without_image = len(dedup_rows) - rows_with_image
    material_filled = sum(1 for r in dedup_rows if r.get("material", "").strip())
    needs_review = sum(1 for r in dedup_rows if (r.get("status") or "").strip() == "needs_review")
    placeholders_lines = coverage["target_lines_as_placeholder"]

    run_report = {
        "generated_at": datetime.now().isoformat(),
        "mode": "merge_scambio_into_final",
        "source_base_catalog": str(args.base_catalog),
        "source_scambio_catalog": str(args.scambio_catalog),
        "scambio_validation": scambio_check,
        "total_rows": len(dedup_rows),
        "producers_count": len(producers),
        **coverage,
        "rows_with_image": rows_with_image,
        "rows_without_image": rows_without_image,
        "material_filled_count": material_filled,
        "needs_review_count": needs_review,
        "duplicates_removed": stats.duplicates_removed,
        "bad_images_removed": stats.bad_images_removed,
        "rows_removed_forbidden_model_name": stats.model_rows_removed,
        "rows_model_name_cleaned": stats.model_rows_cleaned,
        "producer_rows": {row["producer"]: row["rows_count"] for row in producer_summary},
        "warnings": [],
    }
    if "Все модели" in {r.get("target_line", "") for r in dedup_rows if r.get("producer", "") == "АСД"}:
        run_report["warnings"].append("ASD target_line includes 'Все модели' (unexpected)")
    if len(producers) != 10:
        run_report["warnings"].append(f"Expected 10 producers, found {len(producers)}")
    if placeholders_lines > 0:
        run_report["warnings"].append(f"Placeholder target lines present: {placeholders_lines}")

    (args.out / "run_report.json").write_text(json.dumps(run_report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"final_merged ready: {args.out.resolve()}")
    print(json.dumps({
        "total_rows": run_report["total_rows"],
        "producers_count": run_report["producers_count"],
        "rows_without_image": run_report["rows_without_image"],
        "needs_review_count": run_report["needs_review_count"],
        "material_filled_count": run_report["material_filled_count"],
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
