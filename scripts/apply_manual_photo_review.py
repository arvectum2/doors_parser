from __future__ import annotations

import argparse
import csv
import json
import re
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook, Workbook
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


def normalize(text: str) -> str:
    text = (text or "").lower().replace("ё", "е")
    text = re.sub(r"[^a-zа-я0-9]+", " ", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip()


def append_comment(comment: str, addon: str) -> str:
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
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


def read_manual_workbook(path: Path) -> dict[tuple[str, str, str, str], dict]:
    wb = load_workbook(path, data_only=True)
    out: dict[tuple[str, str, str, str], dict] = {}
    for ws in wb.worksheets:
        headers = [str(c.value or "").strip() for c in ws[1]]
        idx = {h: i for i, h in enumerate(headers)}
        required = ["producer", "target_line", "model_name", "source_url", "selected_image_url", "decision", "reviewer_comment"]
        if not all(r in idx for r in required):
            continue
        for row in ws.iter_rows(min_row=2, values_only=True):
            producer = str(row[idx["producer"]] or "").strip()
            target_line = str(row[idx["target_line"]] or "").strip()
            model_name = str(row[idx["model_name"]] or "").strip()
            source_url = str(row[idx["source_url"]] or "").strip()
            if not producer or not model_name:
                continue
            rec = {
                "selected_image_url": str(row[idx["selected_image_url"]] or "").strip(),
                "decision": str(row[idx["decision"]] or "").strip(),
                "reviewer_comment": str(row[idx["reviewer_comment"]] or "").strip(),
            }
            # include candidate urls if available
            for i in range(1, 6):
                k = f"candidate_{i}_url"
                if k in idx:
                    rec[k] = str(row[idx[k]] or "").strip()
            key = (producer, target_line, model_name, source_url)
            out[key] = rec
    return out


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, help="output/final_photo_fix_manual_rules_v2/doors_catalog.csv")
    parser.add_argument("--review", required=True, help="output/manual_photo_review/manual_photo_review.xlsx")
    parser.add_argument("--out", required=True, help="output/final_photo_fix_manual_selected")
    args = parser.parse_args()

    input_path = Path(args.input)
    review_path = Path(args.review)
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    check_dir = out_dir / "check_by_producer"
    check_dir.mkdir(parents=True, exist_ok=True)

    rows = [{k: (r.get(k, "") or "").strip() for k in CATALOG_FIELDS} for r in read_csv(input_path)]
    manual = read_manual_workbook(review_path)

    changed = Counter()
    reparse_rows = []

    for row in rows:
        key = (row.get("producer", ""), row.get("target_line", ""), row.get("model_name", ""), row.get("source_url", ""))
        if key not in manual:
            continue
        m = manual[key]
        old_img = row.get("image_url", "")
        old_status = row.get("status", "")
        old_comment = row.get("comment", "")
        decision = (m.get("decision", "") or "").strip().lower()
        sel_manual = (m.get("selected_image_url", "") or "").strip()

        new_img = old_img
        new_status = old_status
        new_comment = old_comment
        reason = "no_manual_decision"
        action = "kept_old_image"

        if sel_manual:
            new_img = sel_manual
            new_status = "found"
            new_comment = append_comment(new_comment, "manual_selected_image")
            new_comment = append_comment(new_comment, m.get("reviewer_comment", ""))
            reason = "manual_selected_image"
            if new_img != old_img:
                action = "replaced_image"
        elif decision == "use_current":
            new_img = old_img
            reason = "manual_use_current"
            action = "kept_old_image"
        elif decision.startswith("use_candidate_"):
            m_idx = re.search(r"use_candidate_(\d+)", decision)
            if m_idx:
                c_url = m.get(f"candidate_{m_idx.group(1)}_url", "")
                if c_url:
                    new_img = c_url
                    new_status = "found"
                    new_comment = append_comment(new_comment, "manual_use_candidate")
                    new_comment = append_comment(new_comment, m.get("reviewer_comment", ""))
                    reason = f"manual_use_candidate_{m_idx.group(1)}"
                    if new_img != old_img:
                        action = "replaced_image"
                else:
                    new_img = ""
                    new_status = "needs_review"
                    new_comment = append_comment(new_comment, "manual_candidate_missing")
                    reason = "manual_candidate_missing"
                    action = "removed_bad_image" if old_img else "image_not_found"
        elif decision == "clear_image":
            new_img = ""
            new_status = "needs_review"
            new_comment = append_comment(new_comment, "manual_clear_image")
            new_comment = append_comment(new_comment, m.get("reviewer_comment", ""))
            reason = "manual_clear_image"
            action = "removed_bad_image" if old_img else "image_not_found"
        elif decision == "keep_needs_review":
            new_status = "needs_review"
            new_comment = append_comment(new_comment, "manual_keep_needs_review")
            new_comment = append_comment(new_comment, m.get("reviewer_comment", ""))
            reason = "manual_keep_needs_review"
            action = "kept_old_image"

        row["image_url"] = new_img
        row["status"] = new_status
        row["comment"] = new_comment

        if old_img != new_img or old_status != new_status or old_comment != new_comment:
            changed[row.get("producer", "")] += 1

        reparse_rows.append(
            {
                "producer": row.get("producer", ""),
                "target_line": row.get("target_line", ""),
                "model_name": row.get("model_name", ""),
                "source_url": row.get("source_url", ""),
                "old_image_url": old_img,
                "new_image_url": new_img,
                "action": action,
                "reason": reason,
                "http_status": "",
                "content_type": "",
                "width": "",
                "height": "",
                "status_after": new_status,
                "comment_after": new_comment,
            }
        )

    rows_sorted = sorted(rows, key=lambda r: (normalize(r.get("producer", "")), normalize(r.get("target_line", "")), normalize(r.get("model_name", ""))))
    write_csv(out_dir / "doors_catalog.csv", rows_sorted, CATALOG_FIELDS)
    write_xlsx(out_dir / "doors_catalog.xlsx", rows_sorted, CATALOG_FIELDS)

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
    write_csv(out_dir / "wordpress_import.csv", wp_rows, WP_FIELDS)
    write_csv(out_dir / "image_reparse_report.csv", reparse_rows, REPARSE_FIELDS)

    # simple producer slices
    producer_file = {"Олимп": "olymp", "АСД": "asd", "Энтро": "entro", "Хаусдорс": "hausdoors"}
    by_prod: dict[str, list[dict]] = defaultdict(list)
    for r in rows_sorted:
        if r.get("producer", "") in producer_file:
            by_prod[r["producer"]].append(r)
    for p, fname in producer_file.items():
        write_csv(check_dir / f"{fname}.csv", by_prod.get(p, []), CATALOG_FIELDS)
        write_xlsx(check_dir / f"{fname}.xlsx", by_prod.get(p, []), CATALOG_FIELDS, title=fname)

    run_report = {
        "generated_at": datetime.now().isoformat(),
        "mode": "apply_manual_photo_review",
        "input_catalog": str(input_path),
        "review_workbook": str(review_path),
        "total_rows": len(rows_sorted),
        "producers_count": len({r.get("producer", "") for r in rows_sorted if r.get("producer", "")}),
        "rows_with_image": sum(1 for r in rows_sorted if r.get("image_url", "")),
        "rows_without_image": sum(1 for r in rows_sorted if not r.get("image_url", "")),
        "needs_review_count": sum(1 for r in rows_sorted if r.get("status") == "needs_review"),
        "changed_rows_by_producer": dict(changed),
    }
    (out_dir / "run_report.json").write_text(json.dumps(run_report, ensure_ascii=False, indent=2), encoding="utf-8")
    print("apply_manual_photo_review completed")
    print(json.dumps(run_report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

