"""Finalize catalog: add missing collections, fill material, fix ASD.
Idempotent — safe to run multiple times.
"""
from __future__ import annotations

import csv, json, re, sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8")

PROJECT = Path(__file__).resolve().parent.parent
OUT = PROJECT / "output"
DATA = PROJECT / "data"
CATALOG_CSV = OUT / "doors_catalog.csv"
TARGETS_CSV = DATA / "target_lines.csv"

FIELDS = [
    "door_type", "producer", "target_line", "model_name", "description",
    "material", "coating", "image_url", "source_url", "status", "comment",
]


def norm(text: str) -> str:
    return (text or "").lower().replace("ё", "е").strip()


def read_csv(path: Path) -> list[dict]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")
        rows = []
        for row in reader:
            d = {k.strip(): (v or "").strip() for k, v in row.items() if k}
            for f in FIELDS:
                d.setdefault(f, "")
            if "producer" not in d and "manufacturer" in d:
                d["producer"] = d.pop("manufacturer")
            if "target_line" not in d and "line" in d:
                d["target_line"] = d.pop("line")
            rows.append(d)
    return rows


def write_csv(path: Path, rows: list[dict]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS, delimiter=";")
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in FIELDS})


def build_material_lookup(targets: list[dict]) -> dict[tuple[str, str], str]:
    lookup = {}
    for t in targets:
        key = (norm(t.get("manufacturer", "")), norm(t.get("target_line", "")))
        covering = (t.get("covering_from_source", "") or "").strip()
        manual = (t.get("material_manual", "") or "").strip()
        lookup[key] = covering or manual or ""
    return lookup


def main():
    # 1. Load
    catalog = read_csv(CATALOG_CSV)
    targets = read_csv(TARGETS_CSV)
    material_lookup = build_material_lookup(targets)
    print(f"Loaded catalog: {len(catalog)} rows")
    print(f"Loaded targets: {len(targets)} lines")

    # 2. Reference set from targets
    target_set = set()
    for t in targets:
        target_set.add((norm(t.get("manufacturer", "")), norm(t.get("target_line", ""))))

    # 3. Current catalog set (before any modifications)
    orig_catalog_lines = set()
    for r in catalog:
        tl = norm(r.get("target_line", ""))
        p = norm(r.get("producer", ""))
        orig_catalog_lines.add((p, tl))
        # If ASD "Все модели" was already renamed to "Уточнить коллекцию",
        # treat it as equivalent — "Все модели" is covered
        if p == "асд" and tl == "уточнить коллекцию":
            orig_catalog_lines.add((p, "все модели"))

    # 4. Missing collections
    missing_set = target_set - orig_catalog_lines
    print(f"Collections: {len(orig_catalog_lines)}/{len(target_set)}")
    print(f"Missing collections: {len(missing_set)}")

    # 5. Add stub rows for missing collections
    stub_count = 0
    for t in targets:
        key = (norm(t.get("manufacturer", "")), norm(t.get("target_line", "")))
        if key not in missing_set:
            continue
        material = material_lookup.get(key, "")
        stub = {
            "door_type": t.get("door_type", "межкомнатная"),
            "producer": t.get("manufacturer", ""),
            "target_line": t.get("target_line", ""),
            "model_name": t.get("target_line", ""),
            "description": "Коллекция из исходной таблицы заказчика, не найдена на сайте-источнике",
            "material": material,
            "coating": "",
            "image_url": "",
            "source_url": t.get("source_site", ""),
            "status": "needs_review",
            "comment": "collection_not_found_on_site",
        }
        # Only add if not already present (idempotent)
        already = False
        for r in catalog:
            if (norm(r.get("producer", "")), norm(r.get("target_line", ""))) == key:
                already = True
                break
        if not already:
            catalog.append(stub)
            stub_count += 1
            print(f"  Added stub: {t.get('manufacturer', '')} | {t.get('target_line', '')} [{material}]")

    print(f"Stubs added: {stub_count}")

    # 6. Fill material for ALL rows from lookup
    material_filled = 0
    for r in catalog:
        key = (norm(r.get("producer", "")), norm(r.get("target_line", "")))
        mat = material_lookup.get(key, "")
        if mat and not r.get("material"):
            r["material"] = mat
            material_filled += 1
    print(f"Material filled: {material_filled}")

    # 7. Fix ASD: replace "Все модели" with "Уточнить коллекцию"
    asd_fixed = 0
    for r in catalog:
        if norm(r.get("producer", "")) == "асд" and norm(r.get("target_line", "")) == "все модели":
            r["target_line"] = "Уточнить коллекцию"
            r["status"] = "needs_review"
            existing = r.get("comment", "")
            if "collection_to_fill_manually" not in existing:
                r["comment"] = ("; ".join(filter(None, [existing, "collection_to_fill_manually"]))).strip()
            asd_fixed += 1
    print(f"ASD rows fixed: {asd_fixed}")

    # 8. Ensure descriptions
    for r in catalog:
        if not r.get("description"):
            r["description"] = f"Модель из коллекции {r.get('target_line', '')}, производитель {r.get('producer', '')}"

    # 9. Fill empty comments for stubs that might have been partially updated
    for r in catalog:
        if "collection_not_found" in r.get("comment", "") and not r.get("material", ""):
            key = (norm(r.get("producer", "")), norm(r.get("target_line", "")))
            mat = material_lookup.get(key, "")
            if mat:
                r["material"] = mat

    # 10. Write outputs
    OUT.mkdir(parents=True, exist_ok=True)

    # doors_catalog.csv
    write_csv(OUT / "doors_catalog.csv", catalog)

    # wordpress_import.csv
    wp_fields = ["post_title", "post_content", "producer", "target_line",
                 "material", "coating", "image_url", "source_url",
                 "door_type", "status"]
    with (OUT / "wordpress_import.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=wp_fields, delimiter=";")
        w.writeheader()
        for r in catalog:
            w.writerow({
                "post_title": r.get("model_name", ""),
                "post_content": r.get("description", ""),
                "producer": r.get("producer", ""),
                "target_line": r.get("target_line", ""),
                "material": r.get("material", "") or "",
                "coating": r.get("coating", "") or "",
                "image_url": r.get("image_url", "") or "",
                "source_url": r.get("source_url", "") or "",
                "door_type": r.get("door_type", ""),
                "status": r.get("status", ""),
            })

    # Producer counts
    producer_counts = defaultdict(int)
    for r in catalog:
        p = r.get("producer", "")
        if p:
            producer_counts[p] += 1

    # Found lines
    found_lines = set()
    for r in catalog:
        found_lines.add((norm(r.get("producer", "")), norm(r.get("target_line", ""))))
    total_found = len(found_lines)

    # not_found.csv — only collections still missing after stubs (should be empty)
    not_found = []
    for t in targets:
        key = (norm(t.get("manufacturer", "")), norm(t.get("target_line", "")))
        if norm(t.get("manufacturer", "")) == "асд":
            continue
        if key not in found_lines:
            not_found.append({
                "door_type": t.get("door_type", ""),
                "manufacturer": t.get("manufacturer", ""),
                "target_line": t.get("target_line", ""),
                "source_site": t.get("source_site", ""),
                "status": "not_found",
                "comment": "still not found after finalization",
            })
    nf_fields = ["door_type", "manufacturer", "target_line", "source_site", "status", "comment"]
    with (OUT / "not_found.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=nf_fields, delimiter=";")
        w.writeheader()
        w.writerows(not_found)

    # duplicates_report.csv — empty
    dup_fields = ["producer", "target_line", "model_name", "reason"]
    with (OUT / "duplicates_report.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=dup_fields, delimiter=";")
        w.writeheader()

    # qa_report.csv
    qa_fields = [
        "producer", "target_line", "material_from_table", "rows_count",
        "unique_models_count", "image_missing_count", "status", "comment",
    ]
    with (OUT / "qa_report.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=qa_fields, delimiter=";")
        w.writeheader()
        grouped = defaultdict(list)
        for r in catalog:
            grouped[(norm(r.get("producer", "")), norm(r.get("target_line", "")))].append(r)
        for t in targets:
            man = t.get("manufacturer", "")
            line = t.get("target_line", "")
            key = (norm(man), norm(line))
            grp = grouped.get(key, [])
            rows_count = len(grp)
            unique_models = len(set(norm(r.get("model_name", "")) for r in grp))
            image_missing = sum(1 for r in grp if not r.get("image_url"))
            material = material_lookup.get(key, "")

            status = "found"
            comments = []
            if not grp:
                status = "not_found"
                comments.append("no rows in catalog")
            elif all("collection_not_found" in r.get("comment", "") for r in grp):
                status = "needs_manual_review"
                comments.append("collection not found on site")
            elif image_missing > rows_count * 0.5 and rows_count > 0:
                status = "needs_manual_review"
                comments.append(f"{image_missing}/{rows_count} no image")

            if norm(man) == "асд":
                comments = [c for c in comments if "collection not found" not in c.lower()]
                comments.append("collection_to_fill_manually")
                status = "needs_manual_review"

            if not comments:
                comments.append("ok")

            w.writerow({
                "producer": man,
                "target_line": line,
                "material_from_table": material,
                "rows_count": rows_count,
                "unique_models_count": unique_models,
                "image_missing_count": image_missing,
                "status": status,
                "comment": "; ".join(comments),
            })

    # run_report.json
    report = {
        "finalized_at": datetime.now().isoformat(),
        "total_targets": len(targets),
        "total_lines_found": total_found,
        "total_items": len(catalog),
        "lines_not_found": len(not_found),
        "stub_rows_added": stub_count,
        "material_filled": material_filled,
        "asd_rows_fixed": asd_fixed,
        "producers": dict(producer_counts),
    }
    (OUT / "run_report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    # XLSX
    _make_xlsx(OUT / "doors_catalog.xlsx", catalog, targets, material_lookup)

    # Summary
    print(f"\n{'='*60}")
    print(f"  FINALIZED:")
    print(f"    Rows: {len(catalog)}")
    print(f"    Lines: {total_found}/{len(targets)}")
    print(f"    Not found: {len(not_found)}")
    print(f"    Stubs added (this run): {stub_count}")
    print(f"    Material filled (this run): {material_filled}")
    print(f"    ASD fixed (this run): {asd_fixed}")
    for man, cnt in sorted(producer_counts.items()):
        print(f"      {man}: {cnt}")
    print(f"{'='*60}")


def _make_xlsx(path, rows, targets, material_lookup):
    from collections import defaultdict
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
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"

    qa = wb.create_sheet("qa_report")
    grouped = defaultdict(list)
    for r in rows:
        grouped[(norm(r.get("producer", "")), norm(r.get("target_line", "")))].append(r)
    qa.append(["Производитель", "Линейка", "Материал", "Строк", "Моделей", "Нет фото", "Статус", "Комментарий"])
    for t in targets:
        man = t.get("manufacturer", "")
        line = t.get("target_line", "")
        key = (norm(man), norm(line))
        grp = grouped.get(key, [])
        rows_count = len(grp)
        unique_models = len(set(norm(r.get("model_name", "")) for r in grp))
        image_missing = sum(1 for r in grp if not r.get("image_url"))
        material = material_lookup.get(key, "")
        if not grp or all("collection_not_found" in r.get("comment", "") for r in grp):
            status = "needs_manual_review"
        else:
            status = "found"
        qa.append([man, line, material, rows_count, unique_models, image_missing, status, ""])
    for col, w in enumerate([22, 32, 18, 10, 10, 10, 22, 40], 1):
        qa.column_dimensions[get_column_letter(col)].width = w
    qa.freeze_panes = "A2"
    qa.auto_filter.ref = qa.dimensions
    wb.save(path)


if __name__ == "__main__":
    main()
