from __future__ import annotations

import csv
import json
import shutil
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parent.parent
BASELINE_DIR = ROOT / "output" / "final_ready_v11"
FIXED_DIR = ROOT / "output" / "final_ready_v13"
OUT_DIR = ROOT / "output" / "final_hybrid_client_ready"
CHECK_DIR = OUT_DIR / "check_by_producer"
TARGETS_CSV = ROOT / "data" / "target_lines.csv"

OLD_SIX = ("Скамбио Порте", "Лорд", "Аэлита", "Комфорт", "Двери регионов", "Термодур")
NEW_FOUR = ("АСД", "Энтро", "Хаусдорс", "Олимп")

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
REPEATED_FIELDS = [
    "producer",
    "target_line",
    "image_url",
    "repeated_count",
    "affected_model_names",
    "action",
    "reason",
]
UNCHANGED_FIELDS = [
    "producer",
    "baseline_rows",
    "hybrid_rows",
    "exact_match",
    "changed_rows_count",
    "comment",
]
MERGE_REPORT_FIELDS = [
    "producer",
    "source_used",
    "rows_count",
    "changed_vs_baseline",
    "comment",
]


def norm(text: str) -> str:
    return " ".join((text or "").lower().replace("ё", "е").split())


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


def row_key(row: dict) -> tuple:
    return tuple((row.get(k, "") or "").strip() for k in CATALOG_FIELDS)


def producer_rows(rows: list[dict], producer: str) -> list[dict]:
    return [r for r in rows if r.get("producer") == producer]


def recalc_qa(rows: list[dict], target_rows: list[dict]) -> tuple[list[dict], list[dict], dict]:
    grouped: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for row in rows:
        grouped[(norm(row.get("producer", "")), norm(row.get("target_line", "")))].append(row)
    qa_rows, not_found_rows = [], []
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
        nr = sum(1 for r in sub if r.get("status", "") == "needs_review")
        uniq = len({norm(r.get("model_name", "")) for r in sub if r.get("model_name", "")})
        placeholders = [r for r in sub if "collection_not_found_on_site" in (r.get("comment") or "").lower()]
        real = [r for r in sub if r not in placeholders]
        if not sub:
            st = "not_found"
            cm = "checked_catalog_no_rows_for_target_line"
            lnf += 1
            not_found_rows.append(
                {
                    "door_type": t.get("door_type", ""),
                    "manufacturer": p,
                    "target_line": l,
                    "source_site": t.get("source_site", ""),
                    "status": "not_found",
                    "comment": cm,
                }
            )
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
    return qa_rows, not_found_rows, coverage


def main() -> None:
    baseline_catalog = BASELINE_DIR / "doors_catalog.csv"
    fixed_catalog = FIXED_DIR / "doors_catalog.csv"
    if not baseline_catalog.exists():
        raise FileNotFoundError(f"Baseline catalog not found: {baseline_catalog}")
    if not fixed_catalog.exists():
        raise FileNotFoundError(f"Fixed catalog not found: {fixed_catalog}")

    baseline_rows = [{k: r.get(k, "") for k in CATALOG_FIELDS} for r in read_csv(baseline_catalog)]
    fixed_rows = [{k: r.get(k, "") for k in CATALOG_FIELDS} for r in read_csv(fixed_catalog)]
    targets = read_csv(TARGETS_CSV)

    if OUT_DIR.exists():
        shutil.rmtree(OUT_DIR)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    CHECK_DIR.mkdir(parents=True, exist_ok=True)

    old_six_rows = [r for r in baseline_rows if r.get("producer") in OLD_SIX]
    new_four_rows = [r for r in fixed_rows if r.get("producer") in NEW_FOUR]
    hybrid_rows = old_six_rows + new_four_rows

    unchanged_report = []
    all_exact = True
    for p in OLD_SIX:
        b = producer_rows(baseline_rows, p)
        h = producer_rows(hybrid_rows, p)
        b_set = Counter(row_key(r) for r in b)
        h_set = Counter(row_key(r) for r in h)
        exact = b_set == h_set
        all_exact = all_exact and exact
        changed_rows = (b_set - h_set).total() + (h_set - b_set).total()
        unchanged_report.append(
            {
                "producer": p,
                "baseline_rows": len(b),
                "hybrid_rows": len(h),
                "exact_match": "true" if exact else "false",
                "changed_rows_count": changed_rows,
                "comment": "ok" if exact else "mismatch_vs_baseline",
            }
        )

    write_csv(OUT_DIR / "unchanged_baseline_check.csv", unchanged_report, UNCHANGED_FIELDS)

    merge_report = []
    for p in NEW_FOUR:
        b = producer_rows(baseline_rows, p)
        h = producer_rows(hybrid_rows, p)
        changed = Counter(row_key(r) for r in b) != Counter(row_key(r) for r in h)
        merge_report.append(
            {
                "producer": p,
                "source_used": "final_ready_v13",
                "rows_count": len(h),
                "changed_vs_baseline": "allowed",
                "comment": "updated_problem_producer_rows" if changed else "taken_from_fixed_catalog_no_diff_vs_baseline",
            }
        )
    for p in OLD_SIX:
        h = producer_rows(hybrid_rows, p)
        merge_report.append(
            {
                "producer": p,
                "source_used": "baseline_client_sent",
                "rows_count": len(h),
                "changed_vs_baseline": "false",
                "comment": "unchanged_from_baseline",
            }
        )
    write_csv(OUT_DIR / "hybrid_merge_report.csv", merge_report, MERGE_REPORT_FIELDS)

    if not all_exact:
        raise RuntimeError("Baseline exact_match failed for one or more of old six producers. Archive/build aborted.")

    # Save hybrid rows
    hybrid_rows = sorted(hybrid_rows, key=lambda r: (norm(r.get("producer", "")), norm(r.get("target_line", "")), norm(r.get("model_name", ""))))
    write_csv(OUT_DIR / "doors_catalog.csv", hybrid_rows, CATALOG_FIELDS)
    write_xlsx(OUT_DIR / "doors_catalog.xlsx", hybrid_rows, CATALOG_FIELDS)

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
        for r in hybrid_rows
    ]
    write_csv(OUT_DIR / "wordpress_import.csv", wp_rows, WP_FIELDS)

    qa_rows, not_found_rows, coverage = recalc_qa(hybrid_rows, targets)
    write_csv(OUT_DIR / "qa_report.csv", qa_rows, QA_FIELDS)
    write_csv(OUT_DIR / "not_found.csv", not_found_rows, NOT_FOUND_FIELDS)

    # image_reparse_report: changes vs baseline for NEW_FOUR only.
    baseline_map = {(r.get("producer", ""), r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", "")): r for r in baseline_rows if r.get("producer") in NEW_FOUR}
    hybrid_map = {(r.get("producer", ""), r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", "")): r for r in hybrid_rows if r.get("producer") in NEW_FOUR}
    reparse = []
    for k in sorted(set(baseline_map) | set(hybrid_map)):
        b = baseline_map.get(k)
        h = hybrid_map.get(k)
        if b and not h:
            reparse.append(
                {
                    "producer": b.get("producer", ""),
                    "target_line": b.get("target_line", ""),
                    "model_name": b.get("model_name", ""),
                    "source_url": b.get("source_url", ""),
                    "old_image_url": b.get("image_url", ""),
                    "new_image_url": "",
                    "action": "deleted_row",
                    "reason": "replaced_by_fixed_catalog",
                    "http_status": "",
                    "content_type": "",
                    "width": "",
                    "height": "",
                    "status_after": "",
                    "comment_after": "",
                }
            )
        elif h and not b:
            reparse.append(
                {
                    "producer": h.get("producer", ""),
                    "target_line": h.get("target_line", ""),
                    "model_name": h.get("model_name", ""),
                    "source_url": h.get("source_url", ""),
                    "old_image_url": "",
                    "new_image_url": h.get("image_url", ""),
                    "action": "added_row",
                    "reason": "taken_from_fixed_catalog",
                    "http_status": "",
                    "content_type": "",
                    "width": "",
                    "height": "",
                    "status_after": h.get("status", ""),
                    "comment_after": h.get("comment", ""),
                }
            )
        else:
            if any((b.get(f, "") != h.get(f, "")) for f in ("image_url", "status", "comment", "model_name", "target_line", "material", "source_url")):
                reparse.append(
                    {
                        "producer": h.get("producer", ""),
                        "target_line": h.get("target_line", ""),
                        "model_name": h.get("model_name", ""),
                        "source_url": h.get("source_url", ""),
                        "old_image_url": b.get("image_url", ""),
                        "new_image_url": h.get("image_url", ""),
                        "action": "updated_row",
                        "reason": "taken_from_fixed_catalog",
                        "http_status": "",
                        "content_type": "",
                        "width": "",
                        "height": "",
                        "status_after": h.get("status", ""),
                        "comment_after": h.get("comment", ""),
                    }
                )
    write_csv(OUT_DIR / "image_reparse_report.csv", reparse, REPARSE_FIELDS)

    # repeated_image_review from v13 (for touched producers) as context.
    v13_repeat = FIXED_DIR / "repeated_image_review.csv"
    if v13_repeat.exists():
        rpt = [r for r in read_csv(v13_repeat) if r.get("producer") in NEW_FOUR]
    else:
        rpt = []
    write_csv(OUT_DIR / "repeated_image_review.csv", rpt, REPEATED_FIELDS)

    # check_by_producer
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
    producers = sorted({r.get("producer", "") for r in hybrid_rows if r.get("producer", "")}, key=norm)
    for p in producers:
        sub = producer_rows(hybrid_rows, p)
        stem = producer_file.get(p, norm(p).replace(" ", "_"))
        write_csv(CHECK_DIR / f"{stem}.csv", sub, CATALOG_FIELDS)
        write_xlsx(CHECK_DIR / f"{stem}.xlsx", sub, CATALOG_FIELDS, title=stem)

    # duplicate key check
    dup_counter = Counter((r.get("producer", ""), r.get("target_line", ""), r.get("model_name", "")) for r in hybrid_rows)
    dup_count = sum(c - 1 for c in dup_counter.values() if c > 1)

    needs_by_producer = dict(sorted(Counter(r.get("producer", "") for r in hybrid_rows if r.get("status", "") == "needs_review").items()))
    rows_by_producer = dict(sorted(Counter(r.get("producer", "") for r in hybrid_rows if r.get("producer", "")).items()))
    run_report = {
        "generated_at": datetime.now().isoformat(),
        "mode": "final_hybrid_client_ready",
        "baseline_used": str(BASELINE_DIR),
        "fixed_used": str(FIXED_DIR),
        "total_rows": len(hybrid_rows),
        "producers_count": len(producers),
        "rows_by_producer": rows_by_producer,
        "rows_with_image": sum(1 for r in hybrid_rows if r.get("image_url", "")),
        "rows_without_image": sum(1 for r in hybrid_rows if not r.get("image_url", "")),
        "needs_review_count": sum(1 for r in hybrid_rows if r.get("status", "") == "needs_review"),
        "needs_review_by_producer": needs_by_producer,
        "duplicate_count_by_key_producer_line_model": dup_count,
        "old_six_exact_match": True,
        **coverage,
    }
    (OUT_DIR / "run_report.json").write_text(json.dumps(run_report, ensure_ascii=False, indent=2), encoding="utf-8")

    print(json.dumps({"baseline": str(BASELINE_DIR), "total_rows": len(hybrid_rows), "duplicate_count": dup_count}, ensure_ascii=False))


if __name__ == "__main__":
    main()
