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
BASE_DIR = ROOT / "output" / "final_ready_v3"
BASE_CSV = BASE_DIR / "doors_catalog.csv"
TARGETS_CSV = ROOT / "data" / "target_lines.csv"
OUT_DIR = ROOT / "output" / "final_ready_v4"
CHECK_DIR = OUT_DIR / "check_by_producer"

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


def norm(text: str) -> str:
    text = (text or "").lower().replace("ё", "е")
    text = re.sub(r"[^a-zа-я0-9]+", " ", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip()


def norm_model(text: str) -> str:
    text = norm(text)
    text = re.sub(r"\b(дверь|model|модель|variant|вариант|фото)\b", " ", text)
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
    low = (url or "").lower()
    if not low:
        return True
    bad = (
        "promo",
        "resize/20x",
        "/resize/20x/",
        "banner",
        "logo",
        "icon",
        "placeholder",
        "no-image",
        "no_image",
        "open.png",
    )
    return any(x in low for x in bad)


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

    base_rows = [{k: r.get(k, "") for k in CATALOG_FIELDS} for r in read_csv(BASE_CSV)]
    targets = read_csv(TARGETS_CSV)

    olymp_rows = [r for r in base_rows if r.get("producer") == "Олимп"]
    other_rows = [r for r in base_rows if r.get("producer") != "Олимп"]

    img_counts = Counter(r.get("image_url", "") for r in olymp_rows if r.get("image_url", ""))
    repeated_unverified_image_cleared = 0
    olymp_new = []
    for row in olymp_rows:
        r = dict(row)
        is_unverified = "no_visible_caption_photo_preserved_for_manual_naming" in (r.get("comment") or "")
        img = r.get("image_url", "")
        if is_unverified and img and img_counts.get(img, 0) > 3:
            r["image_url"] = ""
            r["status"] = "needs_review"
            r["comment"] = append_comment(r.get("comment", ""), "repeated_unverified_image_cleared")
            repeated_unverified_image_cleared += 1
        olymp_new.append(r)

    rows = other_rows + olymp_new
    rows_sorted = sorted(rows, key=lambda r: (norm(r.get("producer", "")), norm(r.get("target_line", "")), norm(r.get("model_name", ""))))

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

    # minimal rebuild + image reparse reports
    olymp_old = olymp_rows
    olymp_after = [r for r in rows_sorted if r.get("producer") == "Олимп"]
    rebuild_rows = [
        {
            "producer": "Олимп",
            "old_rows": len(olymp_old),
            "new_rows": len(olymp_after),
            "old_variant_rows": 0,
            "new_variant_rows": 0,
            "real_model_rows": sum(1 for r in olymp_after if r.get("status") == "found"),
            "generated_model_rows": sum(1 for r in olymp_after if "no_visible_caption_photo_preserved_for_manual_naming" in (r.get("comment") or "")),
            "rows_with_image": sum(1 for r in olymp_after if r.get("image_url", "")),
            "rows_without_image": sum(1 for r in olymp_after if not r.get("image_url", "")),
            "repeated_image_gt_10_count": 0,
            "needs_review_count": sum(1 for r in olymp_after if r.get("status") == "needs_review"),
            "comment": "repeated_unverified_image_cleanup",
        }
    ]
    write_csv(OUT_DIR / "rebuild_report.csv", rebuild_rows, [
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
    ])

    reparse_rows = []
    old_map = {(r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", "")): r for r in olymp_old}
    new_map = {(r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", "")): r for r in olymp_after}
    for key in sorted(set(old_map) | set(new_map)):
        o = old_map.get(key)
        n = new_map.get(key)
        if not o or not n:
            continue
        if o.get("image_url", "") == n.get("image_url", "") and o.get("status", "") == n.get("status", "") and o.get("comment", "") == n.get("comment", ""):
            continue
        reparse_rows.append(
            {
                "producer": "Олимп",
                "target_line": n.get("target_line", ""),
                "model_name": n.get("model_name", ""),
                "source_url": n.get("source_url", ""),
                "old_image_url": o.get("image_url", ""),
                "new_image_url": n.get("image_url", ""),
                "action": "updated_row",
                "reason": "repeated_unverified_image_cleared",
                "http_status": "",
                "content_type": "",
                "width": "",
                "height": "",
                "status_after": n.get("status", ""),
                "comment_after": n.get("comment", ""),
            }
        )
    write_csv(OUT_DIR / "image_reparse_report.csv", reparse_rows, [
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
    ])

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
    for p in producers:
        sub = [r for r in rows_sorted if r.get("producer") == p]
        stem = producer_file.get(p, norm(p).replace(" ", "_"))
        write_csv(CHECK_DIR / f"{stem}.csv", sub, CATALOG_FIELDS)
        write_xlsx(CHECK_DIR / f"{stem}.xlsx", sub, CATALOG_FIELDS, title=stem)

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
        "mode": "final_ready_v4_olymp_repeated_image_cleanup",
        "source_catalog": str(BASE_CSV),
        "old_total_rows": len(base_rows),
        "total_rows": len(rows_sorted),
        "producers_count": len(producers),
        **coverage,
        "rows_with_image": sum(1 for r in rows_sorted if r.get("image_url", "")),
        "rows_without_image": sum(1 for r in rows_sorted if not r.get("image_url", "")),
        "needs_review_count": sum(1 for r in rows_sorted if r.get("status") == "needs_review"),
        "rows_by_producer": dict(sorted(Counter(r.get("producer", "") for r in rows_sorted if r.get("producer", "")).items())),
        "olymp_rows": len(olymp_after),
        "olymp_rows_with_image": sum(1 for r in olymp_after if r.get("image_url", "")),
        "olymp_rows_without_image": sum(1 for r in olymp_after if not r.get("image_url", "")),
        "repeated_unverified_image_cleared_count": repeated_unverified_image_cleared,
        "unchanged_producers_check": unchanged,
    }
    (OUT_DIR / "run_report.json").write_text(json.dumps(run_report, ensure_ascii=False, indent=2), encoding="utf-8")

    print("Built output/final_ready_v4")
    print(json.dumps({"total_rows": len(rows_sorted), "olymp_rows": len(olymp_after), "cleared": repeated_unverified_image_cleared}, ensure_ascii=False))


if __name__ == "__main__":
    main()

