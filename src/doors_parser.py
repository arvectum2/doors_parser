from __future__ import annotations

import argparse
import csv
import json
import re
import sys
import traceback
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from collections import Counter
from urllib.parse import urlparse

# Ensure project root is on sys.path so "from src.adapters" works
_project_root = Path(__file__).resolve().parent.parent
if str(_project_root) not in sys.path:
    sys.path.insert(0, str(_project_root))

import requests
import yaml
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from src.adapters import (
    TargetLine, DoorItem, normalize, now_str, fetch, dedup_items,
    is_generic_page, BAD_IMAGE_KEYWORDS, get_adapter,
    normalize_dver_model_name,
)


# ── Config loading ────────────────────────────────────────────────────────────

def detect_csv_dialect(sample: str) -> csv.Dialect:
    first_line = sample.splitlines()[0] if sample.splitlines() else ""
    if ";" in first_line:
        class SemiDialect(csv.excel):
            delimiter = ";"
        return SemiDialect
    try:
        return csv.Sniffer().sniff(sample, delimiters=";,\t")
    except csv.Error:
        class SemiDialect(csv.excel):
            delimiter = ";"
        return SemiDialect


def read_targets(path: Path) -> list[TargetLine]:
    if not path.exists():
        raise FileNotFoundError(f"target file not found: {path}")
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        sample = f.read(8192)
        f.seek(0)
        dialect = detect_csv_dialect(sample)
        reader = csv.DictReader(f, dialect=dialect)
        rows = []
        for row in reader:
            row = {str(k).strip(): (v or "").strip() for k, v in row.items() if k}
            if not row.get("manufacturer") or not row.get("target_line"):
                continue
            rows.append(TargetLine(
                door_type=row.get("door_type", ""),
                manufacturer=row.get("manufacturer", ""),
                target_line=row.get("target_line", ""),
                material_manual="",
                covering_from_source="",
                source_site=row.get("source_site", ""),
                source_note=row.get("source_note", ""),
            ))
        if not rows:
            raise RuntimeError(
                "target_lines.csv read but no rows found. "
                "Check headers: door_type;manufacturer;target_line;..."
            )
        print(f"  Loaded target lines: {len(rows)}")
        return rows


def load_sites(path: Path) -> dict[str, dict]:
    data = yaml.safe_load(path.read_text(encoding="utf-8"))
    sites = {}
    for s in data["sites"]:
        key = normalize(s["manufacturer"])
        s.setdefault("seeds", [s["base_url"]])
        s.setdefault("adapter", "generic")
        sites[key] = s
    return sites


# ── Output writers ────────────────────────────────────────────────────────────

def write_csv(path: Path, rows: list[DoorItem]) -> None:
    fields = ["door_type", "producer", "target_line", "model_name", "description",
              "material", "coating", "image_url", "source_url", "status", "comment"]
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields, delimiter=";")
        w.writeheader()
        for r in rows:
            d = asdict(r)
            d["producer"] = d.pop("manufacturer", "")
            d["target_line"] = d.pop("line", "")
            w.writerow({k: d.get(k, "") for k in fields})


def write_wp_csv(path: Path, rows: list[DoorItem]) -> None:
    fields = ["post_title", "post_content", "producer", "line",
              "material", "coating", "image_url", "source_url",
              "door_type", "status"]
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields, delimiter=";")
        w.writeheader()
        for r in rows:
            w.writerow({
                "post_title": r.model_name,
                "post_content": r.description,
                "producer": r.manufacturer,
                "line": r.line,
                "material": r.material,
                "coating": r.coating,
                "image_url": r.image_url,
                "source_url": r.source_url,
                "door_type": r.door_type,
                "status": r.status,
            })


def write_not_found(path: Path, targets: list[TargetLine], rows: list[DoorItem]) -> None:
    found = {(normalize(r.manufacturer), normalize(r.line)) for r in rows}
    reasons = {
        "Профайн": "page_not_found — URL не найден на сайте",
        "Дэлия": "page_not_found — URL не найден на сайте",
        "Экошпон": "line_found_but_no_models — материал, не линейка (нет страницы)",
    }
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        fields = ["door_type", "manufacturer", "target_line", "source_site", "status", "comment"]
        w = csv.DictWriter(f, fieldnames=fields, delimiter=";")
        w.writeheader()
        for t in targets:
            key = (normalize(t.manufacturer), normalize(t.target_line))
            if key not in found:
                reason = reasons.get(t.target_line, "needs_manual_review")
                w.writerow({
                    "door_type": t.door_type,
                    "manufacturer": t.manufacturer,
                    "target_line": t.target_line,
                    "source_site": t.source_site,
                    "status": "not_found",
                    "comment": reason,
                })


def write_xlsx(path: Path, rows: list[DoorItem], targets: list[TargetLine], report: dict) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "parsed_catalog"
    headers = ["Тип двери", "Производитель", "Линейка", "Название", "Описание",
               "Материал", "Покрытие", "Ссылка на фото", "Ссылка на карточку",
               "Статус", "Комментарий"]
    ws.append(headers)
    for r in rows:
        ws.append([r.door_type, r.manufacturer, r.line, r.model_name,
                   r.description, r.material, r.coating, r.image_url,
                   r.source_url, r.status, r.comment])
    widths = [16, 22, 28, 40, 70, 18, 16, 60, 60, 14, 40]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    nf = wb.create_sheet("not_found")
    nf.append(["Тип двери", "Производитель", "Линейка", "Сайт", "Статус", "Комментарий"])
    found = {(normalize(r.manufacturer), normalize(r.line)) for r in rows}
    for t in targets:
        if (normalize(t.manufacturer), normalize(t.target_line)) not in found:
            nf.append([t.door_type, t.manufacturer, t.target_line, t.source_site,
                       "not_found", "проверить вручную / дописать селектор"])
    for cell in nf[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="9C6500")
    for col, w in enumerate([16, 24, 32, 45, 18, 45], 1):
        nf.column_dimensions[get_column_letter(col)].width = w
    nf.freeze_panes = "A2"
    nf.auto_filter.ref = nf.dimensions

    rp = wb.create_sheet("run_report")
    rp.append(["Показатель", "Значение"])
    rp.append(["Дата", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    rp.append(["Целевых линеек", report.get("total_targets", len(targets))])
    rp.append(["Найдено строк", len(rows)])
    rp.append(["Не найдено линеек", nf.max_row - 1])
    if rows:
        rp.append(["С моделями(status=found)", sum(1 for r in rows if r.status == "found")])
        rp.append(["С моделями(status=needs_review)", sum(1 for r in rows if r.status == "needs_review")])
    rp.append(["Комментарий", "Парсятся только линейки из target_lines.csv. Материал/покрытие пустые."])
    rp.append([])
    rp.append(["Производитель", "Страниц", "Моделей", "Ошибки"])
    for man, info in report.get("sites", {}).items():
        errs = len(info.get("errors", []))
        rp.append([man, info.get("pages_processed", 0), info.get("models_found", 0), errs])
    if report.get("errors"):
        rp.append([])
        rp.append(["Глобальные ошибки", ""])
        for e in report["errors"]:
            rp.append([e, ""])
    for cell in rp[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="548235")
    rp.column_dimensions["A"].width = 32
    rp.column_dimensions["B"].width = 60
    rp.column_dimensions["C"].width = 12
    rp.column_dimensions["D"].width = 12
    wb.save(path)


# ── QA report ─────────────────────────────────────────────────────────────────

def write_qa_report(path: Path, all_rows: list[DoorItem], targets: list[TargetLine],
                    report: dict, site_rows: dict[str, list[DoorItem]]) -> None:
    grouped: dict[tuple[str, str], list[DoorItem]] = {}
    for r in all_rows:
        grouped.setdefault((normalize(r.manufacturer), normalize(r.line)), []).append(r)

    # Track which source URLs appear in many lines
    url_line_count: Counter = Counter()
    for r in all_rows:
        url_line_count[(r.source_url, normalize(r.line))] += 1
    url_total_lines: Counter = Counter()
    for (url, line), cnt in url_line_count.items():
        url_total_lines[url] += 1

    with path.open("w", encoding="utf-8-sig", newline="") as f:
        fields = [
            "producer", "target_line", "rows_count", "unique_models_count",
            "unique_source_urls_count", "bad_images_count",
            "generic_pages_count", "duplicate_rows_count",
            "needs_review_count", "status", "comment",
        ]
        w = csv.DictWriter(f, fieldnames=fields, delimiter=";")
        w.writeheader()

        for t in targets:
            key = (normalize(t.manufacturer), normalize(t.target_line))
            rows = grouped.get(key, [])
            nf_status = "found" if rows else "not_found"

            rows_count = len(rows)
            unique_models = len(set(normalize(r.model_name) for r in rows))
            unique_urls = len(set(r.source_url for r in rows))
            bad_images = sum(1 for r in rows if any(k in (r.image_url or "").lower() for k in BAD_IMAGE_KEYWORDS))
            generic_pages = sum(1 for r in rows if is_generic_page(r.model_name, r.source_url))
            duplicate_rows = rows_count - unique_models if rows else 0
            needs_review = sum(1 for r in rows if r.status == "needs_review")

            comments = []
            status = nf_status

            if rows_count == 0:
                status = "not_found"
                comments.append("no rows")
            elif generic_pages > rows_count * 0.3:
                status = "needs_manual_review"
                comments.append(f"{generic_pages}/{rows_count} generic pages")
            elif bad_images > rows_count * 0.5:
                status = "needs_manual_review"
                comments.append(f"{bad_images}/{rows_count} bad images")
            elif needs_review > rows_count * 0.5:
                status = "needs_manual_review"
                comments.append(f"{needs_review}/{rows_count} needs_review")

            # Check if source URLs are shared across many lines
            for r in rows:
                if url_total_lines.get(r.source_url, 0) >= 3:
                    status = "needs_manual_review"
                    comments.append(f"url shared across {url_total_lines[r.source_url]} lines")
                    break

            if not comments:
                comments.append("ok")

            w.writerow({
                "producer": t.manufacturer,
                "target_line": t.target_line,
                "rows_count": rows_count,
                "unique_models_count": unique_models,
                "unique_source_urls_count": unique_urls,
                "bad_images_count": bad_images,
                "generic_pages_count": generic_pages,
                "duplicate_rows_count": duplicate_rows,
                "needs_review_count": needs_review,
                "status": status,
                "comment": "; ".join(comments),
            })


def save_intermediate(out_dir: Path, all_rows: list[DoorItem], targets: list[TargetLine],
                      report: dict, site_rows: dict[str, list[DoorItem]], suffix: str = "") -> None:
    prefix = "partial_" if suffix else ""
    write_csv(out_dir / f"{prefix}doors_catalog.csv", all_rows)
    write_wp_csv(out_dir / f"{prefix}wordpress_import.csv", all_rows)
    write_not_found(out_dir / f"{prefix}not_found.csv", targets, all_rows)
    write_qa_report(out_dir / f"{prefix}qa_report.csv", all_rows, targets, report, site_rows)
    (out_dir / f"{prefix}run_report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    # Also save per-producer CSVs for merge_results.py compatibility
    for man_name, man_rows in site_rows.items():
        man_dir = out_dir / normalize(man_name)
        man_dir.mkdir(parents=True, exist_ok=True)
        write_csv(man_dir / "doors_catalog.csv", man_rows)


# ── CLI arg parsing ──────────────────────────────────────────────────────────

def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="Парсер каталога дверей по целевым линейкам заказчика")
    ap.add_argument("--targets", type=Path, default=Path("data/target_lines.csv"))
    ap.add_argument("--sites", type=Path, default=Path("config/sites.yaml"))
    ap.add_argument("--out", type=Path, default=Path("output"))
    ap.add_argument("--only", type=str, default=None, help="Обработать только одного производителя (по имени из target_lines.csv)")
    ap.add_argument("--limit-sites", type=int, default=None, help="Максимум сайтов для обработки")
    ap.add_argument("--max-pages-per-site", type=int, default=200, help="Максимум страниц на сайт")
    ap.add_argument("--debug", action="store_true", help="Подробный вывод")
    ap.add_argument("--resume", action="store_true", help="Дозапись: пропустить уже обработанные сайты")
    ap.add_argument("--postprocess-only", action="store_true",
                    help="Только постобработка (группировка дублей, выбор лучшего фото) без парсинга")
    return ap.parse_args(argv)


# ── Main ─────────────────────────────────────────────────────────────────────

def main() -> None:
    args = parse_args()
    out_dir = args.out
    out_dir.mkdir(parents=True, exist_ok=True)

    print("=" * 60)
    print("  Doors Catalog Parser")
    print("=" * 60)

    if args.postprocess_only:
        from scripts.cleanup_table_images import cleanup

        class PostprocArgs:
            input = out_dir / "doors_catalog.csv"
            out = out_dir
            targets = args.targets
        cleanup(PostprocArgs())
        return

    targets = read_targets(args.targets)
    sites_map = load_sites(args.sites)

    grouped: dict[str, list[TargetLine]] = {}
    for t in targets:
        grouped.setdefault(normalize(t.manufacturer), []).append(t)

    print(f"\nTotal manufacturers: {len(grouped)}")

    completed_sites: set[str] = set()
    if args.resume:
        report_path = out_dir / "run_report.json"
        if report_path.exists():
            try:
                existing = json.loads(report_path.read_text(encoding="utf-8"))
                completed_sites = set(existing.get("completed_sites", []))
                print(f"  Resuming, already completed: {len(completed_sites)} sites")
            except Exception:
                pass

    session = requests.Session()
    all_rows: list[DoorItem] = []
    site_rows: dict[str, list[DoorItem]] = {}
    report: dict = {
        "start_time": datetime.now().isoformat(),
        "total_targets": len(targets),
        "total_lines_found": 0,
        "sites": {},
        "errors": [],
        "completed_sites": [],
    }

    manufacturers_to_process = list(grouped.items())
    if args.only:
        n_only = normalize(args.only)
        manufacturers_to_process = [(k, v) for k, v in manufacturers_to_process if k == n_only]
        if not manufacturers_to_process:
            print(f"[WARN] Manufacturer '{args.only}' not found in targets. Available:")
            for k in grouped:
                print(f"       - {k}")
            sys.exit(1)
        print(f"\n  Filtered to manufacturer: {args.only}")

    if args.limit_sites:
        manufacturers_to_process = manufacturers_to_process[:args.limit_sites]
        print(f"  Limited to {args.limit_sites} site(s)")

    for idx, (man_normalized, tlist) in enumerate(manufacturers_to_process, 1):
        if args.resume and man_normalized in completed_sites:
            print(f"\n[{idx}/{len(manufacturers_to_process)}] SKIP (resume) {tlist[0].manufacturer}")
            continue

        cfg = sites_map.get(man_normalized)
        if not cfg:
            print(f"\n[{idx}/{len(manufacturers_to_process)}] [WARN] No site config for: {tlist[0].manufacturer}")
            report["errors"].append(f"No site config for {tlist[0].manufacturer}")
            continue

        man_name = tlist[0].manufacturer
        door_type_default = cfg.get("door_type", "межкомнатная")

        # Auto-add line_url_map keys as direct URLs
        line_url_map = cfg.get("line_url_map", None)
        if line_url_map:
            direct_list = cfg.get("direct_urls", [])
            for u in line_url_map:
                if u not in direct_list:
                    direct_list.append(u)
            cfg["direct_urls"] = direct_list

        print(f"\n[{idx}/{len(manufacturers_to_process)}] {man_name}")
        print(f"      Site: {cfg['base_url']}")
        print(f"      Target lines: {len(tlist)}")
        print(f"      Adapter: {cfg.get('adapter', 'generic')}")

        try:
            adapter_cls = get_adapter(cfg.get("adapter", "generic"))
            adapter = adapter_cls(
                session, cfg, tlist, door_type_default,
                max_pages=args.max_pages_per_site,
                debug=args.debug,
            )
            rows, site_log = adapter.run()
            all_rows.extend(rows)
            site_rows[man_name] = rows

            site_log["items"] = len(rows)
            site_log["manufacturer"] = man_name
            report["sites"][man_name] = site_log
            report["completed_sites"].append(man_normalized)

            # Count statuses for progress
            found_count = sum(1 for r in rows if r.status == "found")
            review_count = sum(1 for r in rows if r.status == "needs_review")
            print(f"      found={found_count} needs_review={review_count}")

            save_intermediate(out_dir, all_rows, targets, report, site_rows, suffix="partial")

            print(f"      Total models so far: {len(all_rows)}")

        except Exception as e:
            err = traceback.format_exc()
            (out_dir / f"error_{man_normalized}.log").write_text(err, encoding="utf-8")
            report["errors"].append(f"{man_name}: {e}")
            print(f"      [ERROR] {man_name}: {e}")

    # Final global dedup (producer + target_line + model_name + image_url + source_url)
    seen: set[tuple[str, str, str, str, str]] = set()
    deduped = []
    for r in all_rows:
        key = (normalize(r.manufacturer), normalize(r.line), normalize(r.model_name),
               normalize(r.image_url), normalize(r.source_url))
        if key in seen:
            continue
        seen.add(key)
        deduped.append(r)
    all_rows = deduped

    # Clean dver.com model names (remove sizes, "Распродажа")
    for r in all_rows:
        if "dver.com" in r.source_url or "скамбио" in normalize(r.manufacturer):
            cleaned, orig = normalize_dver_model_name(r.model_name)
            if cleaned != orig:
                if r.comment:
                    r.comment = f"raw_name: {orig}; " + r.comment
                else:
                    r.comment = f"raw_name: {orig}"
                r.model_name = cleaned

    report["total_items"] = len(all_rows)
    report["end_time"] = datetime.now().isoformat()

    found_lines = set()
    for r in all_rows:
        found_lines.add((normalize(r.manufacturer), normalize(r.line)))
    report["total_lines_found"] = len(found_lines)

    print(f"\n{'=' * 60}")
    print(f"  Writing output files...")
    write_csv(out_dir / "doors_catalog.csv", all_rows)
    write_wp_csv(out_dir / "wordpress_import.csv", all_rows)
    write_not_found(out_dir / "not_found.csv", targets, all_rows)
    write_qa_report(out_dir / "qa_report.csv", all_rows, targets, report, site_rows)
    write_xlsx(out_dir / "doors_catalog.xlsx", all_rows, targets, report)
    (out_dir / "run_report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    # Clean up partials
    for f in out_dir.glob("partial_*"):
        f.unlink()

    # Summary with status breakdown
    found_total = sum(1 for r in all_rows if r.status == "found")
    review_total = sum(1 for r in all_rows if r.status == "needs_review")
    nf_count = report["total_targets"] - report["total_lines_found"]

    print(f"\n  DONE:")
    print(f"    Lines: {report['total_lines_found']}/{report['total_targets']} ({report['total_lines_found']/report['total_targets']*100:.1f}%)")
    print(f"    Models found: {found_total}")
    print(f"    Needs review: {review_total}")
    print(f"    Not found lines: {nf_count}")
    print(f"    Errors: {len(report['errors'])}")
    print(f"    Output: {out_dir.resolve()}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
