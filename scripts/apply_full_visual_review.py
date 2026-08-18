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
INPUT_CSV = ROOT / "output" / "final_ready_v11" / "doors_catalog.csv"
DECISIONS_CSV = ROOT / "output" / "final_ready_v11" / "full_visual_review" / "manual_full_door_decisions.csv"
TARGETS_CSV = ROOT / "data" / "target_lines.csv"
OUT_DIR = ROOT / "output" / "final_ready_v12"
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
REPORT_FIELDS = [
    "producer",
    "target_line",
    "model_name",
    "source_url",
    "current_image_url",
    "decision",
    "selected_image_url",
    "action_applied",
    "status_after",
    "comment_after",
]


def norm(text: str) -> str:
    text = (text or "").lower().replace("ё", "е")
    text = re.sub(r"[^a-zа-я0-9]+", " ", text, flags=re.I)
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


def recalc_qa(rows: list[dict], target_rows: list[dict]) -> tuple[list[dict], dict]:
    grouped: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for row in rows:
        grouped[(norm(row.get("producer", "")), norm(row.get("target_line", "")))].append(row)
    qa_rows = []
    lt = lr = lp = lnf = 0
    for t in target_rows:
        p = (t.get("manufacturer") or "").strip()
        l = (t.get("target_line") or "").strip()
        if not p or not l:
            continue
        lt += 1
        key = (norm(p), norm(l))
        sub = grouped.get(key, [])
        if not sub and norm(p) == norm("АСД") and norm(l) == norm("Все модели"):
            sub = grouped.get((norm("АСД"), norm("Уточнить коллекцию")), [])
        with_img = sum(1 for r in sub if r.get("image_url", ""))
        miss = len(sub) - with_img
        nr = sum(1 for r in sub if r.get("status") == "needs_review")
        uniq = len({norm(r.get("model_name", "")) for r in sub if r.get("model_name", "")})
        placeholders = [r for r in sub if "collection_not_found_on_site" in (r.get("comment") or "").lower()]
        real = [r for r in sub if r not in placeholders]
        if not sub:
            st = "not_found"
            cm = "checked_catalog_no_rows_for_target_line"
            lnf += 1
        elif real:
            lr += 1
            st = "needs_review" if (miss or nr) else "found"
            cm = f"with_image={with_img}; missing_image={miss}; needs_review={nr}"
        else:
            lp += 1
            st = "needs_review"
            cm = "collection_not_found_on_site"
        qa_rows.append(
            {
                "producer": p,
                "target_line": l,
                "expected_material": (t.get("covering_from_source") or t.get("material_manual") or "").strip(),
                "rows_count": len(sub),
                "unique_models_count": uniq,
                "with_image_count": with_img,
                "missing_image_count": miss,
                "needs_review_count": nr,
                "bad_image_count": 0,
                "status": st,
                "comment": cm,
            }
        )
    coverage = {
        "target_lines_total": lt,
        "target_lines_with_real_models": lr,
        "target_lines_as_placeholder": lp,
        "target_lines_not_found": lnf,
    }
    return qa_rows, coverage


def main() -> None:
    if OUT_DIR.exists():
        shutil.rmtree(OUT_DIR)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    CHECK_DIR.mkdir(parents=True, exist_ok=True)

    rows = [{k: r.get(k, "") for k in CATALOG_FIELDS} for r in read_csv(INPUT_CSV)]
    decisions = read_csv(DECISIONS_CSV)
    targets = read_csv(TARGETS_CSV)

    decision_map = {}
    for d in decisions:
        k = (
            d.get("producer", ""),
            d.get("target_line", ""),
            d.get("model_name", ""),
            d.get("source_url", ""),
            d.get("current_image_url", ""),
        )
        decision_map[k] = d

    out_rows: list[dict] = []
    report_rows: list[dict] = []
    stats = Counter()
    for row in rows:
        k = (
            row.get("producer", ""),
            row.get("target_line", ""),
            row.get("model_name", ""),
            row.get("source_url", ""),
            row.get("image_url", ""),
        )
        d = decision_map.get(k)
        if not d:
            out_rows.append(row)
            continue
        dec = (d.get("decision") or "").strip().lower()
        selected = (d.get("selected_image_url") or "").strip()
        action = "unreviewed"
        new_row = dict(row)
        if dec == "keep":
            action = "keep"
            stats["keep"] += 1
        elif dec == "clear":
            new_row["image_url"] = ""
            new_row["status"] = "needs_review"
            new_row["comment"] = append_comment(new_row.get("comment", ""), "manual_visual_clear_not_full_door")
            action = "clear"
            stats["clear"] += 1
        elif dec == "replace":
            if selected:
                new_row["image_url"] = selected
                new_row["status"] = "found"
                new_row["comment"] = append_comment(new_row.get("comment", ""), "manual_visual_replaced_full_door")
                action = "replace"
                stats["replace"] += 1
            else:
                action = "replace_missing_selected_image_url"
                stats["replace_missing_selected_image_url"] += 1
        elif dec == "delete_row":
            action = "delete_row"
            stats["delete_row"] += 1
            report_rows.append(
                {
                    "producer": d.get("producer", ""),
                    "target_line": d.get("target_line", ""),
                    "model_name": d.get("model_name", ""),
                    "source_url": d.get("source_url", ""),
                    "current_image_url": d.get("current_image_url", ""),
                    "decision": d.get("decision", ""),
                    "selected_image_url": d.get("selected_image_url", ""),
                    "action_applied": action,
                    "status_after": "",
                    "comment_after": "",
                }
            )
            continue
        else:
            action = "unreviewed"
            stats["unreviewed"] += 1

        out_rows.append(new_row)
        report_rows.append(
            {
                "producer": d.get("producer", ""),
                "target_line": d.get("target_line", ""),
                "model_name": d.get("model_name", ""),
                "source_url": d.get("source_url", ""),
                "current_image_url": d.get("current_image_url", ""),
                "decision": d.get("decision", ""),
                "selected_image_url": d.get("selected_image_url", ""),
                "action_applied": action,
                "status_after": new_row.get("status", ""),
                "comment_after": new_row.get("comment", ""),
            }
        )

    out_rows = sorted(out_rows, key=lambda r: (norm(r.get("producer", "")), norm(r.get("target_line", "")), norm(r.get("model_name", ""))))
    write_csv(OUT_DIR / "doors_catalog.csv", out_rows, CATALOG_FIELDS)
    write_xlsx(OUT_DIR / "doors_catalog.xlsx", out_rows, CATALOG_FIELDS)

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
        for r in out_rows
    ]
    write_csv(OUT_DIR / "wordpress_import.csv", wp_rows, WP_FIELDS)

    qa_rows, coverage = recalc_qa(out_rows, targets)
    write_csv(OUT_DIR / "qa_report.csv", qa_rows, QA_FIELDS)
    write_csv(OUT_DIR / "full_visual_review_report.csv", report_rows, REPORT_FIELDS)

    # Re-emit image_reparse style report for compatibility from review report.
    write_csv(OUT_DIR / "image_reparse_report.csv", report_rows, REPORT_FIELDS)

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
    producers = sorted({r.get("producer", "") for r in out_rows if r.get("producer", "")}, key=norm)
    for p in producers:
        sub = [r for r in out_rows if r.get("producer") == p]
        stem = producer_file.get(p, norm(p).replace(" ", "_"))
        write_csv(CHECK_DIR / f"{stem}.csv", sub, CATALOG_FIELDS)
        write_xlsx(CHECK_DIR / f"{stem}.xlsx", sub, CATALOG_FIELDS, title=stem)

    run_report = {
        "generated_at": datetime.now().isoformat(),
        "mode": "final_ready_v12_from_manual_visual_review",
        "source_catalog": str(INPUT_CSV),
        "source_decisions": str(DECISIONS_CSV),
        "total_rows": len(out_rows),
        "producers_count": len(producers),
        "rows_by_producer": dict(sorted(Counter(r.get("producer", "") for r in out_rows if r.get("producer", "")).items())),
        "rows_with_image": sum(1 for r in out_rows if r.get("image_url", "")),
        "rows_without_image": sum(1 for r in out_rows if not r.get("image_url", "")),
        "needs_review_count": sum(1 for r in out_rows if r.get("status", "") == "needs_review"),
        **coverage,
        "manual_decision_stats": dict(stats),
    }
    (OUT_DIR / "run_report.json").write_text(json.dumps(run_report, ensure_ascii=False, indent=2), encoding="utf-8")
    print("Applied manual visual review to output/final_ready_v12")


if __name__ == "__main__":
    main()

