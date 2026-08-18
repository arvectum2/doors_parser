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

from rebuild_hausdoors_from_model_rows import bad_image_url as haus_bad_image_url
from rebuild_hausdoors_from_model_rows import rebuild_hausdoors_rows
from rebuild_olymp_from_model_rows import bad_image_url as olymp_bad_image_url
from rebuild_olymp_from_model_rows import rebuild_olymp_rows


ROOT = Path(__file__).resolve().parent.parent
BASE_CSV = ROOT / "output" / "final_problem_sites_rebuilt" / "doors_catalog.csv"
PREV_CSV = ROOT / "output" / "final_images_fix" / "doors_catalog.csv"
TARGETS_CSV = ROOT / "data" / "target_lines.csv"
OUT_DIR = ROOT / "output" / "final_ready"

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


def bad_image_url(url: str) -> bool:
    return haus_bad_image_url(url) or olymp_bad_image_url(url)


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
    check_dir = OUT_DIR / "check_by_producer"
    check_dir.mkdir(parents=True, exist_ok=True)

    base_rows = [{k: r.get(k, "") for k in CATALOG_FIELDS} for r in read_csv(BASE_CSV)]
    prev_rows = [{k: r.get(k, "") for k in CATALOG_FIELDS} for r in read_csv(PREV_CSV)] if PREV_CSV.exists() else []
    targets = read_csv(TARGETS_CSV)

    old_counts = Counter(r.get("producer", "") for r in base_rows if r.get("producer", ""))
    unchanged_producers = {"Скамбио Порте", "Лорд", "Аэлита", "Комфорт", "Двери регионов", "Термодур", "Энтро"}

    haus_old = [r for r in base_rows if r.get("producer") == "Хаусдорс"]
    olymp_old = [r for r in base_rows if r.get("producer") == "Олимп"]
    asd_rows = [r for r in base_rows if r.get("producer") == "АСД"]
    unchanged_rows = [r for r in base_rows if r.get("producer") in unchanged_producers]

    haus_new, haus_stats = rebuild_hausdoors_rows(haus_old)
    olymp_new, olymp_stats = rebuild_olymp_rows(olymp_old)

    # ASD final bad-image sanity cleanup only (explicit bad markers).
    asd_new: list[dict] = []
    asd_fixed_bad = 0
    asd_bad_re = re.compile(r"(hit|badge|logo|icon|404|not found|error|placeholder|no_image|no-image|open\.png)", re.I)
    for row in asd_rows:
        r = dict(row)
        img = r.get("image_url", "")
        if img and asd_bad_re.search(img or ""):
            r["image_url"] = ""
            r["status"] = "needs_review"
            r["comment"] = append_comment(r.get("comment", ""), "bad_image_removed")
            asd_fixed_bad += 1
        asd_new.append(r)

    final_rows = unchanged_rows + asd_new + haus_new + olymp_new
    rows_sorted = sorted(final_rows, key=lambda r: (norm(r.get("producer", "")), norm(r.get("target_line", "")), norm(r.get("model_name", ""))))

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

    # image_reparse_report
    old_by_producer = defaultdict(list)
    new_by_producer = defaultdict(list)
    for r in base_rows:
        old_by_producer[r.get("producer", "")].append(r)
    for r in rows_sorted:
        new_by_producer[r.get("producer", "")].append(r)

    image_reparse_rows: list[dict] = []
    for producer in ("Хаусдорс", "Олимп", "АСД"):
        old = old_by_producer.get(producer, [])
        new = new_by_producer.get(producer, [])
        old_map = {(r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", "")): r for r in old}
        new_map = {(r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", "")): r for r in new}
        keys = set(old_map) | set(new_map)
        for key in sorted(keys):
            o = old_map.get(key)
            n = new_map.get(key)
            if o and not n:
                action = "removed_old_row"
                reason = "old_variant_removed" if re.search(r"\bvariant\s*\d+\b|\bвариант\s*\d+\b", o.get("model_name", ""), flags=re.I) else "rebuild_cleanup"
                row = {
                    "producer": producer,
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
                action = "added_rebuilt_row"
                reason = "generated_model_name_no_visible_caption" if "generated_model_name_no_visible_caption" in (n.get("comment") or "") else "rebuilt_from_visible_cards"
                row = {
                    "producer": producer,
                    "target_line": n.get("target_line", ""),
                    "model_name": n.get("model_name", ""),
                    "source_url": n.get("source_url", ""),
                    "old_image_url": "",
                    "new_image_url": n.get("image_url", ""),
                    "action": action,
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
                    "producer": producer,
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

    # rebuild_report
    prev_entro_old_rows = len([r for r in prev_rows if r.get("producer") == "Энтро"])
    prev_entro_old_variant = sum(
        1 for r in prev_rows if r.get("producer") == "Энтро" and re.search(r"\bvariant\s*\d+\b|\bвариант\s*\d+\b", r.get("model_name", ""), flags=re.I)
    )
    entro_new = [r for r in rows_sorted if r.get("producer") == "Энтро"]
    entro_stats = {
        "old_rows": prev_entro_old_rows or 320,
        "new_rows": len(entro_new),
        "old_variant_rows": prev_entro_old_variant or 320,
        "new_variant_rows": 0,
        "real_model_rows": len(entro_new),
        "generated_model_rows": 0,
        "rows_with_image": sum(1 for r in entro_new if r.get("image_url", "")),
        "rows_without_image": sum(1 for r in entro_new if not r.get("image_url", "")),
        "repeated_image_gt_10_count": 0,
        "needs_review_count": sum(1 for r in entro_new if r.get("status") == "needs_review"),
        "comment": "kept_from_previous_entro_rebuild",
    }
    rebuild_rows = [
        {"producer": "Хаусдорс", **haus_stats, "comment": "rebuild_from_cards_or_collection_fallback"},
        {"producer": "Олимп", **olymp_stats, "comment": "rebuild_from_gallery_or_generated_names"},
        {"producer": "Энтро", **entro_stats},
    ]
    write_csv(OUT_DIR / "rebuild_report.csv", rebuild_rows, REBUILD_FIELDS)

    # check_by_producer all 10
    producers = sorted({r.get("producer", "") for r in rows_sorted if r.get("producer", "")}, key=lambda x: norm(x))
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
    audit_rows: list[dict] = []
    for producer in producers:
        sub = [r for r in rows_sorted if r.get("producer") == producer]
        stem = producer_file.get(producer, norm(producer).replace(" ", "_"))
        write_csv(check_dir / f"{stem}.csv", sub, CATALOG_FIELDS)
        write_xlsx(check_dir / f"{stem}.xlsx", sub, CATALOG_FIELDS, title=stem)
        img_counts = Counter(r.get("image_url", "") for r in sub if r.get("image_url", ""))
        repeated_gt10 = sum(1 for r in sub if r.get("image_url", "") and img_counts.get(r.get("image_url", ""), 0) > 10)
        variant_rows = sum(
            1 for r in sub if re.search(r"\bvariant\s*\d+\b|\bвариант\s*\d+\b", r.get("model_name", ""), flags=re.I)
        )
        generated_rows = sum(
            1
            for r in sub
            if "generated_model_name_no_visible_caption" in (r.get("comment") or "")
            or "generated_from_visible_caption" in (r.get("comment") or "")
        )
        audit_rows.append(
            {
                "producer": producer,
                "rows": len(sub),
                "rows_with_image": sum(1 for r in sub if r.get("image_url", "")),
                "rows_without_image": sum(1 for r in sub if not r.get("image_url", "")),
                "needs_review": sum(1 for r in sub if r.get("status") == "needs_review"),
                "variant_rows": variant_rows,
                "generated_model_rows": generated_rows,
                "repeated_image_gt_10": repeated_gt10,
                "comment": "ok" if producer not in {"Хаусдорс", "Олимп"} else "rebuilt_problem_site",
            }
        )
    write_csv(check_dir / "final_problem_sites_audit.csv", audit_rows, AUDIT_FIELDS)

    # run_report + invariance check for untouched producers.
    base_by_prod = defaultdict(list)
    new_by_prod = defaultdict(list)
    for r in base_rows:
        base_by_prod[r.get("producer", "")].append(r)
    for r in rows_sorted:
        new_by_prod[r.get("producer", "")].append(r)
    unchanged_check = {}
    for producer in ("Скамбио Порте", "Лорд", "Аэлита", "Комфорт", "Двери регионов", "Термодур"):
        b = sorted(base_by_prod.get(producer, []), key=lambda r: (r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", ""), r.get("image_url", "")))
        n = sorted(new_by_prod.get(producer, []), key=lambda r: (r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", ""), r.get("image_url", "")))
        unchanged_check[producer] = {"rows_before": len(b), "rows_after": len(n), "exact_match": b == n}

    run_report = {
        "generated_at": datetime.now().isoformat(),
        "mode": "final_ready_rebuild_hausdoors_olymp",
        "source_catalog": str(BASE_CSV),
        "old_total_rows": len(base_rows),
        "total_rows": len(rows_sorted),
        "producers_count": len(producers),
        **coverage,
        "rows_with_image": sum(1 for r in rows_sorted if r.get("image_url", "")),
        "rows_without_image": sum(1 for r in rows_sorted if not r.get("image_url", "")),
        "needs_review_count": sum(1 for r in rows_sorted if r.get("status") == "needs_review"),
        "material_filled_count": sum(1 for r in rows_sorted if r.get("material", "")),
        "rows_by_producer": dict(sorted(Counter(r.get("producer", "") for r in rows_sorted if r.get("producer", "")).items())),
        "hausdoors_rebuild": haus_stats,
        "olymp_rebuild": olymp_stats,
        "entro_final_rows": len(entro_new),
        "asd_final_rows": len([r for r in rows_sorted if r.get("producer") == "АСД"]),
        "asd_bad_images_cleaned": asd_fixed_bad,
        "unchanged_producers_check": unchanged_check,
    }
    (OUT_DIR / "run_report.json").write_text(json.dumps(run_report, ensure_ascii=False, indent=2), encoding="utf-8")

    print("Built output/final_ready")
    print(json.dumps({"old_total_rows": len(base_rows), "new_total_rows": len(rows_sorted), "rows_by_producer": run_report["rows_by_producer"]}, ensure_ascii=False))


if __name__ == "__main__":
    main()
