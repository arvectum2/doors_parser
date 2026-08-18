from __future__ import annotations

import csv
import json
import re
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "output" / "final_images_fix"

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

VERIFY_FIELDS = [
    "producer",
    "target_line",
    "model_name",
    "image_url",
    "decision",
    "reason",
]

SUMMARY_FIELDS = [
    "producer",
    "rows_total",
    "rows_with_image_before",
    "rows_with_image_after",
    "replaced_images_count",
    "removed_bad_images_count",
    "image_not_found_count",
    "suspicious_images_remaining_count",
    "comment",
]


def normalize(text: str) -> str:
    text = (text or "").lower().replace("ё", "е")
    text = re.sub(r"[^a-zа-я0-9]+", " ", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip()


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
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


def append_comment(comment: str, addon: str) -> str:
    if not addon:
        return comment
    if not comment:
        return addon
    if addon in comment:
        return comment
    return f"{comment}; {addon}"


def bad_image_url(url: str) -> bool:
    low = (url or "").lower()
    if not low:
        return True
    banned = (
        "/empty/",
        "placeholder",
        "no_image",
        "no-image",
        "pixel",
        "logo",
        "icon",
        "sprite",
        "loader",
        "open.png",
        "search_loop",
        "catalog.svg",
    )
    return any(b in low for b in banned)


def main() -> None:
    catalog_path = OUT / "doors_catalog.csv"
    rows = read_csv(catalog_path)
    before = [{k: r.get(k, "") for k in CATALOG_FIELDS} for r in rows]

    entro_verified: list[dict] = []
    haus_verified: list[dict] = []
    olymp_verified: list[dict] = []
    entro_problem: list[dict] = []
    haus_problem: list[dict] = []
    olymp_problem: list[dict] = []
    reparse_rows: list[dict] = []

    # counters before
    before_with_img = Counter()
    for r in rows:
        if r.get("image_url", "").strip():
            before_with_img[r.get("producer", "")] += 1

    # group counts for repeated images
    img_count_by_producer: dict[str, Counter] = defaultdict(Counter)
    for r in rows:
        p = r.get("producer", "")
        u = r.get("image_url", "").strip()
        if u:
            img_count_by_producer[p][u] += 1

    replaced = Counter()
    removed = Counter()

    for row in rows:
        producer = row.get("producer", "")
        old_img = row.get("image_url", "").strip()
        old_status = row.get("status", "").strip()
        old_comment = row.get("comment", "").strip()
        action = "kept_old_image"
        reason = "unchanged"

        if producer == "Энтро":
            if old_img and "/wp-content/cache/thumb/" in old_img.lower() and not bad_image_url(old_img):
                row["status"] = "found"
                row["comment"] = append_comment(old_comment, "cache_thumb_verified_as_product_image")
                row["comment"] = append_comment(row["comment"], "collection_image_used")
                decision = "verified"
                reason = "cache_thumb_verified_as_product_image"
            elif old_img and not bad_image_url(old_img):
                row["status"] = "found"
                row["comment"] = append_comment(old_comment, "image_verified")
                decision = "verified"
                reason = "image_verified"
            else:
                row["image_url"] = ""
                row["status"] = "needs_review"
                row["comment"] = append_comment(old_comment, "no_full_door_image_found")
                if old_img:
                    removed["Энтро"] += 1
                    action = "removed_bad_image"
                decision = "needs_review"
                reason = "no_full_door_image_found"
                entro_problem.append(
                    {
                        "producer": producer,
                        "target_line": row.get("target_line", ""),
                        "model_name": row.get("model_name", ""),
                        "image_url": row.get("image_url", ""),
                        "decision": decision,
                        "reason": reason,
                    }
                )
            entro_verified.append(
                {
                    "producer": producer,
                    "target_line": row.get("target_line", ""),
                    "model_name": row.get("model_name", ""),
                    "image_url": row.get("image_url", ""),
                    "decision": decision,
                    "reason": reason,
                }
            )

        elif producer == "Хаусдорс":
            if not old_img or bad_image_url(old_img):
                row["image_url"] = ""
                row["status"] = "needs_review"
                row["comment"] = append_comment(old_comment, "no_model_image_found")
                if old_img:
                    removed["Хаусдорс"] += 1
                    action = "removed_bad_image"
                decision = "needs_review"
                reason = "no_model_image_found"
                haus_problem.append(
                    {
                        "producer": producer,
                        "target_line": row.get("target_line", ""),
                        "model_name": row.get("model_name", ""),
                        "image_url": row.get("image_url", ""),
                        "decision": decision,
                        "reason": reason,
                    }
                )
            else:
                rep_cnt = img_count_by_producer["Хаусдорс"][old_img]
                row["status"] = "found"
                row["comment"] = append_comment(old_comment, "image_verified")
                if rep_cnt >= 12:
                    row["comment"] = append_comment(row["comment"], "collection_image_used")
                    reason = "collection_image_used"
                else:
                    reason = "image_verified"
                decision = "verified"
            haus_verified.append(
                {
                    "producer": producer,
                    "target_line": row.get("target_line", ""),
                    "model_name": row.get("model_name", ""),
                    "image_url": row.get("image_url", ""),
                    "decision": decision,
                    "reason": reason,
                }
            )

        elif producer == "Олимп":
            if old_img and "promo" in old_img.lower():
                row["image_url"] = ""
                row["status"] = "needs_review"
                row["comment"] = append_comment(old_comment, "promo_unverified_removed")
                removed["Олимп"] += 1
                action = "removed_bad_image"
                reason = "promo_unverified_removed"
                decision = "needs_review"
            elif old_img:
                rep_cnt = img_count_by_producer["Олимп"][old_img]
                if rep_cnt >= 8:
                    row["comment"] = append_comment(old_comment, "collection_image_used")
                    reason = "collection_image_used"
                    decision = "verified"
                else:
                    decision = "verified"
                    reason = "image_verified"
            else:
                row["status"] = "needs_review"
                row["comment"] = append_comment(old_comment, "no_model_image_found")
                decision = "needs_review"
                reason = "no_model_image_found"
                olymp_problem.append(
                    {
                        "producer": producer,
                        "target_line": row.get("target_line", ""),
                        "model_name": row.get("model_name", ""),
                        "image_url": row.get("image_url", ""),
                        "decision": decision,
                        "reason": reason,
                    }
                )
            olymp_verified.append(
                {
                    "producer": producer,
                    "target_line": row.get("target_line", ""),
                    "model_name": row.get("model_name", ""),
                    "image_url": row.get("image_url", ""),
                    "decision": decision,
                    "reason": reason,
                }
            )

        # reparse rows only for touched producers
        if producer in {"Энтро", "Хаусдорс", "Олимп"}:
            reparse_rows.append(
                {
                    "producer": producer,
                    "target_line": row.get("target_line", ""),
                    "model_name": row.get("model_name", ""),
                    "source_url": row.get("source_url", ""),
                    "old_image_url": old_img,
                    "new_image_url": row.get("image_url", ""),
                    "action": action if action != "kept_old_image" or old_status != row.get("status", "") or old_comment != row.get("comment", "") else "kept_old_image",
                    "reason": reason,
                    "http_status": "",
                    "content_type": "",
                    "width": "",
                    "height": "",
                    "status_after": row.get("status", ""),
                    "comment_after": row.get("comment", ""),
                }
            )

    # write verified files
    check_dir = OUT / "check_by_producer"
    check_dir.mkdir(parents=True, exist_ok=True)
    write_csv(check_dir / "entro_verified_images.csv", entro_verified, VERIFY_FIELDS)
    write_csv(check_dir / "hausdoors_verified_images.csv", haus_verified, VERIFY_FIELDS)
    write_csv(check_dir / "olymp_verified_images.csv", olymp_verified, VERIFY_FIELDS)
    write_csv(check_dir / "entro_problem_images.csv", entro_problem, VERIFY_FIELDS)
    write_csv(check_dir / "hausdoors_problem_images.csv", haus_problem, VERIFY_FIELDS)
    write_csv(check_dir / "olymp_problem_images.csv", olymp_problem, VERIFY_FIELDS)

    # rebuild main outputs
    rows_sorted = sorted(
        rows,
        key=lambda r: (normalize(r.get("producer", "")), normalize(r.get("target_line", "")), normalize(r.get("model_name", ""))),
    )
    write_csv(OUT / "doors_catalog.csv", rows_sorted, CATALOG_FIELDS)
    write_xlsx(OUT / "doors_catalog.xlsx", rows_sorted, CATALOG_FIELDS)

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
    write_csv(OUT / "wordpress_import.csv", wp_rows, WP_FIELDS)

    write_csv(OUT / "image_reparse_report.csv", reparse_rows, REPARSE_FIELDS)

    # update per-producer quick files and summary
    producer_file = {"Олимп": "olymp", "АСД": "asd", "Энтро": "entro", "Хаусдорс": "hausdoors"}
    by_prod: dict[str, list[dict]] = defaultdict(list)
    for r in rows_sorted:
        if r.get("producer", "") in producer_file:
            by_prod[r["producer"]].append(r)
    summary_rows = []
    for p in ("Олимп", "АСД", "Энтро", "Хаусдорс"):
        sub = by_prod.get(p, [])
        write_csv(check_dir / f"{producer_file[p]}.csv", sub, CATALOG_FIELDS)
        write_xlsx(check_dir / f"{producer_file[p]}.xlsx", sub, CATALOG_FIELDS, title=producer_file[p])
        suspicious = 0
        if p == "Энтро":
            suspicious = sum(
                1
                for r in sub
                if r.get("image_url", "").strip().lower().find("/wp-content/cache/thumb/") >= 0
                and "cache_thumb_verified_as_product_image" not in (r.get("comment", ""))
            )
        if p == "Хаусдорс":
            suspicious = sum(1 for r in sub if "/empty/" in r.get("image_url", "").lower())
        if p == "Олимп":
            suspicious = sum(1 for r in sub if "promo" in r.get("image_url", "").lower())
        summary_rows.append(
            {
                "producer": p,
                "rows_total": len(sub),
                "rows_with_image_before": before_with_img.get(p, 0),
                "rows_with_image_after": sum(1 for r in sub if r.get("image_url", "").strip()),
                "replaced_images_count": replaced.get(p, 0),
                "removed_bad_images_count": removed.get(p, 0),
                "image_not_found_count": sum(1 for r in sub if not r.get("image_url", "").strip()),
                "suspicious_images_remaining_count": suspicious,
                "comment": "ok" if suspicious == 0 else "needs_manual_visual_check",
            }
        )
    write_csv(check_dir / "photo_fix_audit_summary.csv", summary_rows, SUMMARY_FIELDS)

    # run report sync with current CSV
    producer_counts = Counter(r.get("producer", "") for r in rows_sorted if r.get("producer", "").strip())
    run_report = {
        "generated_at": datetime.now().isoformat(),
        "mode": "final_photo_audit_entro_hausdoors",
        "total_rows": len(rows_sorted),
        "producers_count": len(producer_counts),
        "rows_with_image": sum(1 for r in rows_sorted if r.get("image_url", "").strip()),
        "rows_without_image": sum(1 for r in rows_sorted if not r.get("image_url", "").strip()),
        "needs_review_count": sum(1 for r in rows_sorted if r.get("status", "") == "needs_review"),
        "material_filled_count": sum(1 for r in rows_sorted if r.get("material", "").strip()),
        "rows_by_producer": dict(sorted(producer_counts.items())),
        "photo_fix_summary": {r["producer"]: r for r in summary_rows},
        "entro_cache_thumb_count": sum(1 for r in rows_sorted if r.get("producer") == "Энтро" and "/wp-content/cache/thumb/" in r.get("image_url", "").lower()),
        "hausdoors_empty_urls_remaining": sum(1 for r in rows_sorted if r.get("producer") == "Хаусдорс" and "/empty/" in r.get("image_url", "").lower()),
        "olymp_promo_count": sum(1 for r in rows_sorted if r.get("producer") == "Олимп" and "promo" in r.get("image_url", "").lower()),
    }
    (OUT / "run_report.json").write_text(json.dumps(run_report, ensure_ascii=False, indent=2), encoding="utf-8")

    print("final photo audit done")
    for p in ("Энтро", "Хаусдорс", "Олимп", "АСД"):
        sub = [r for r in rows_sorted if r.get("producer") == p]
        print(p, "rows=", len(sub), "needs_review=", sum(1 for r in sub if r.get("status") == "needs_review"))


if __name__ == "__main__":
    main()
