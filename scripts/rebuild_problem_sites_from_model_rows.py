from __future__ import annotations

import csv
import json
import re
import shutil
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parent.parent
BASE_DIR = ROOT / "output" / "final_images_fix"
OUT_DIR = ROOT / "output" / "final_problem_sites_rebuilt"
TARGETS_CSV = ROOT / "data" / "target_lines.csv"
ENTRO_CANDIDATES = ROOT / "output" / "image_candidates_v2" / "entro_candidates.csv"

TARGET_PRODUCERS = {"Энтро", "Хаусдорс", "Олимп", "АСД"}

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

IMAGE_REPARSE_FIELDS = [
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
    "new_real_model_rows",
    "rows_with_image",
    "rows_without_image",
    "comment",
]

ENTRO_REBUILT_FIELDS = [
    "target_line",
    "model_name",
    "image_url",
    "source_url",
    "extraction_method",
    "comment",
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
        writer = csv.DictWriter(f, fieldnames=fields, delimiter=";")
        writer.writeheader()
        for row in rows:
            writer.writerow({k: row.get(k, "") for k in fields})


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


def bad_image_url(url: str) -> bool:
    low = (url or "").lower()
    if not low:
        return True
    banned = (
        "logo",
        "icon",
        "badge",
        "sticker",
        "label",
        "sale",
        "hit",
        "placeholder",
        "no_image",
        "no-image",
        "pixel",
        "open.png",
        "search_loop",
        "catalog.svg",
        "sprite",
        "loader",
        "favicon",
        "/empty/",
    )
    return any(b in low for b in banned)


def is_noise_candidate_text(text: str) -> bool:
    t = norm(text)
    if not t:
        return True
    banned_exact = {
        "где купить",
        "купить",
        "главная",
        "каталог",
        "модельный ряд",
        "двери в интерьере",
        "подробнее",
        "характеристики",
        "описание",
        "контакты",
    }
    if t in banned_exact:
        return True
    banned_parts = (
        "где купить",
        "доставка",
        "оплата",
        "гарантия",
        "whatsapp",
        "telegram",
        "vk.com",
    )
    return any(p in t for p in banned_parts)


def tilda_restore(url: str) -> str:
    u = url or ""
    u = re.sub(r"/-/resizeb?/\d+x/", "/", u, flags=re.I)
    u = re.sub(r"/-/cover/\d+x\d+/", "/", u, flags=re.I)
    if "thb.tildacdn.com" in u:
        u = u.replace("thb.tildacdn.com", "static.tildacdn.com")
    return u


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


def rebuild_entro_rows(old_entro_rows: list[dict], candidate_rows: list[dict]) -> tuple[list[dict], list[dict], int]:
    line_defaults: dict[str, dict] = {}
    for row in old_entro_rows:
        line = row.get("target_line", "")
        if line and line not in line_defaults:
            line_defaults[line] = {
                "door_type": row.get("door_type", ""),
                "producer": "Энтро",
                "target_line": line,
                "description": row.get("description", ""),
                "material": row.get("material", ""),
                "coating": row.get("coating", ""),
                "source_url": row.get("source_url", ""),
            }

    prepared: list[dict] = []
    old_variant_rows = 0
    for row in old_entro_rows:
        if re.search(r"\bvariant\s*\d+\b|\bвариант\s*\d+\b", row.get("model_name", ""), flags=re.I):
            old_variant_rows += 1

    for c in candidate_rows:
        if (c.get("producer") or "") != "Энтро":
            continue
        txt = (c.get("candidate_text") or "").strip()
        img = (c.get("candidate_image_url") or "").strip()
        block = (c.get("candidate_block") or "").strip().lower()
        if not txt or not img:
            continue
        if block not in {"model_row", "other"}:
            continue
        if is_noise_candidate_text(txt):
            continue
        if bad_image_url(img):
            continue
        line = (c.get("target_line") or "").strip()
        source_url = (c.get("source_url") or "").strip()
        prepared.append(
            {
                "target_line": line,
                "model_name": re.sub(r"\s+", " ", txt).strip(),
                "image_url": img,
                "source_url": source_url,
                "candidate_block": block,
            }
        )

    unique_rows: list[dict] = []
    seen: set[tuple[str, str, str]] = set()
    for item in prepared:
        key = (norm(item["target_line"]), norm(item["model_name"]), item["image_url"])
        if key in seen:
            continue
        seen.add(key)
        unique_rows.append(item)

    rebuilt_rows: list[dict] = []
    rebuilt_report_rows: list[dict] = []
    for item in sorted(unique_rows, key=lambda x: (norm(x["target_line"]), norm(x["model_name"]))):
        line = item["target_line"]
        defaults = line_defaults.get(line, {})
        desc = defaults.get("description", "")
        if not desc or "variant" in norm(desc):
            desc = f"Модель {item['model_name']} из коллекции {line}, производитель Энтро."

        image_url = item["image_url"]
        status = "found"
        comment = "rebuilt_from_model_row"
        if bad_image_url(image_url):
            image_url = ""
            status = "needs_review"
            comment = append_comment(comment, "no_full_door_image_found")

        row = {
            "door_type": defaults.get("door_type", "межкомнатная"),
            "producer": "Энтро",
            "target_line": line,
            "model_name": item["model_name"],
            "description": desc,
            "material": defaults.get("material", ""),
            "coating": defaults.get("coating", ""),
            "image_url": image_url,
            "source_url": item["source_url"] or defaults.get("source_url", ""),
            "status": status,
            "comment": comment,
        }
        rebuilt_rows.append(row)
        rebuilt_report_rows.append(
            {
                "target_line": line,
                "model_name": item["model_name"],
                "image_url": image_url,
                "source_url": row["source_url"],
                "extraction_method": "candidate_text_model_row",
                "comment": comment,
            }
        )

    return rebuilt_rows, rebuilt_report_rows, old_variant_rows


def main() -> None:
    if OUT_DIR.exists():
        shutil.rmtree(OUT_DIR)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    (OUT_DIR / "check_by_producer").mkdir(parents=True, exist_ok=True)

    base_rows = read_csv(BASE_DIR / "doors_catalog.csv")
    targets = read_csv(TARGETS_CSV)
    entro_candidates = read_csv(ENTRO_CANDIDATES)

    old_counts = Counter(r.get("producer", "") for r in base_rows if r.get("producer", "").strip())
    old_entro_rows = [r for r in base_rows if r.get("producer") == "Энтро"]
    non_entro_rows = [r for r in base_rows if r.get("producer") != "Энтро"]

    rebuilt_entro_rows, entro_rebuilt_report, old_variant_rows = rebuild_entro_rows(old_entro_rows, entro_candidates)

    # Keep other producers as-is, only apply safe cleanup for 3 problem producers.
    image_reparse_rows: list[dict] = []
    refreshed_rows: list[dict] = []
    old_map: dict[tuple[str, str, str, str], dict] = {}
    for r in base_rows:
        key = (r.get("producer", ""), r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", ""))
        old_map[key] = r

    for row in non_entro_rows + rebuilt_entro_rows:
        producer = row.get("producer", "")
        old = old_map.get((producer, row.get("target_line", ""), row.get("model_name", ""), row.get("source_url", "")), {})
        old_img = old.get("image_url", "")
        old_status = old.get("status", "")
        old_comment = old.get("comment", "")
        new_img = row.get("image_url", "")
        new_status = row.get("status", "")
        new_comment = row.get("comment", "")
        action = "kept_old_image"
        reason = "unchanged"

        if producer in {"Хаусдорс", "Олимп"} and new_img:
            restored = tilda_restore(new_img)
            if restored != new_img:
                new_img = restored
                row["image_url"] = restored
                new_comment = append_comment(new_comment, "image_url_restored_from_tilda_resize")
                row["comment"] = new_comment
                reason = "tilda_original_used"
                action = "replaced_image" if old_img != new_img else "kept_old_image"

        if producer in TARGET_PRODUCERS:
            if producer in {"Хаусдорс", "Олимп", "АСД"}:
                if bad_image_url(new_img):
                    row["image_url"] = ""
                    row["status"] = "needs_review"
                    row["comment"] = append_comment(row.get("comment", ""), "no_model_image_found")
                    new_img = ""
                    new_status = "needs_review"
                    new_comment = row["comment"]
                    reason = "bad_or_placeholder_image_removed"
                    action = "removed_bad_image" if old_img else "image_not_found"
                elif not row.get("image_url", ""):
                    row["status"] = "needs_review"
                    row["comment"] = append_comment(row.get("comment", ""), "no_model_image_found")
                    new_status = "needs_review"
                    new_comment = row["comment"]
                    action = "image_not_found"
                    reason = "image_not_found"

            if producer == "Энтро":
                reason = "rebuilt_from_model_row"
                action = "replaced_image"
                if not row.get("image_url", ""):
                    action = "image_not_found"
                    reason = "no_model_image_found"

            image_reparse_rows.append(
                {
                    "producer": producer,
                    "target_line": row.get("target_line", ""),
                    "model_name": row.get("model_name", ""),
                    "source_url": row.get("source_url", ""),
                    "old_image_url": old_img,
                    "new_image_url": row.get("image_url", ""),
                    "action": action,
                    "reason": reason,
                    "http_status": "",
                    "content_type": "",
                    "width": "",
                    "height": "",
                    "status_after": row.get("status", ""),
                    "comment_after": row.get("comment", ""),
                }
            )

        refreshed_rows.append({k: row.get(k, "") for k in CATALOG_FIELDS})

    # Mark large repeated images for Hausdoors/Olymp explicitly.
    for producer in ("Хаусдорс", "Олимп"):
        prod_rows = [r for r in refreshed_rows if r.get("producer") == producer]
        counts = Counter(r.get("image_url", "") for r in prod_rows if r.get("image_url", ""))
        for r in prod_rows:
            img = r.get("image_url", "")
            if img and counts.get(img, 0) > 10:
                r["comment"] = append_comment(r.get("comment", ""), "collection_image_used")
                if producer == "Хаусдорс":
                    r["status"] = "needs_review"

    rows_sorted = sorted(
        refreshed_rows,
        key=lambda r: (norm(r.get("producer", "")), norm(r.get("target_line", "")), norm(r.get("model_name", ""))),
    )

    # Core outputs
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

    write_csv(OUT_DIR / "image_reparse_report.csv", image_reparse_rows, IMAGE_REPARSE_FIELDS)
    write_csv(
        OUT_DIR / "image_audit.csv",
        image_reparse_rows,
        IMAGE_REPARSE_FIELDS,
    )

    # Keep duplicates_report format in output.
    src_dup = BASE_DIR / "duplicates_report.csv"
    if src_dup.exists():
        shutil.copy2(src_dup, OUT_DIR / "duplicates_report.csv")
    else:
        write_csv(
            OUT_DIR / "duplicates_report.csv",
            [],
            [
                "producer",
                "target_line",
                "normalized_model_name",
                "duplicates_count_before",
                "kept_model_name",
                "kept_image_url",
                "removed_image_urls",
                "reason",
            ],
        )

    # Per-producer check slices for 4 problem producers.
    producer_to_file = {"Энтро": "entro", "Хаусдорс": "hausdoors", "Олимп": "olymp", "АСД": "asd"}
    for producer, stem in producer_to_file.items():
        subset = [r for r in rows_sorted if r.get("producer") == producer]
        write_csv(OUT_DIR / "check_by_producer" / f"{stem}.csv", subset, CATALOG_FIELDS)
        write_xlsx(OUT_DIR / "check_by_producer" / f"{stem}.xlsx", subset, CATALOG_FIELDS, title=stem)

    # Rebuild reports.
    new_counts = Counter(r.get("producer", "") for r in rows_sorted if r.get("producer", "").strip())
    rebuild_rows: list[dict] = []
    for producer in ("Энтро", "Хаусдорс", "Олимп", "АСД"):
        old_rows = old_counts.get(producer, 0)
        new_rows = new_counts.get(producer, 0)
        prod_rows = [r for r in rows_sorted if r.get("producer") == producer]
        rebuild_rows.append(
            {
                "producer": producer,
                "old_rows": old_rows,
                "new_rows": new_rows,
                "old_variant_rows": old_variant_rows if producer == "Энтро" else 0,
                "new_real_model_rows": len(prod_rows) if producer == "Энтро" else "",
                "rows_with_image": sum(1 for r in prod_rows if r.get("image_url", "").strip()),
                "rows_without_image": sum(1 for r in prod_rows if not r.get("image_url", "").strip()),
                "comment": "rebuilt_from_model_row_candidates" if producer == "Энтро" else "safe_image_cleanup_only",
            }
        )
    write_csv(OUT_DIR / "rebuild_report.csv", rebuild_rows, REBUILD_FIELDS)
    write_csv(OUT_DIR / "entro_rebuilt_from_model_rows.csv", entro_rebuilt_report, ENTRO_REBUILT_FIELDS)

    rows_with_image = sum(1 for r in rows_sorted if r.get("image_url", "").strip())
    rows_without_image = len(rows_sorted) - rows_with_image
    needs_review = sum(1 for r in rows_sorted if (r.get("status") or "").strip() == "needs_review")
    material_filled = sum(1 for r in rows_sorted if r.get("material", "").strip())

    run_report = {
        "generated_at": datetime.now().isoformat(),
        "mode": "final_problem_sites_rebuilt",
        "source_catalog": str(BASE_DIR / "doors_catalog.csv"),
        "total_rows_before": len(base_rows),
        "total_rows": len(rows_sorted),
        "producers_count": len(new_counts),
        **coverage,
        "rows_with_image": rows_with_image,
        "rows_without_image": rows_without_image,
        "needs_review_count": needs_review,
        "material_filled_count": material_filled,
        "rows_by_producer": dict(sorted(new_counts.items())),
        "problem_sites_rebuild_summary": {
            row["producer"]: row for row in rebuild_rows
        },
    }
    (OUT_DIR / "run_report.json").write_text(json.dumps(run_report, ensure_ascii=False, indent=2), encoding="utf-8")

    print("Done: output/final_problem_sites_rebuilt")
    print(f"old_total_rows={len(base_rows)} new_total_rows={len(rows_sorted)}")
    print(f"old_entro_rows={len(old_entro_rows)} new_entro_rows={len(rebuilt_entro_rows)} old_variant_rows={old_variant_rows}")
    print(f"rows_with_image={rows_with_image} rows_without_image={rows_without_image} needs_review={needs_review}")


if __name__ == "__main__":
    main()
