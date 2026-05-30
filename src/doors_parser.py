from __future__ import annotations

import argparse
import csv
import json
import re
import shutil
import sys
import traceback
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from collections import Counter, defaultdict
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
                material_manual=row.get("material_manual", ""),
                covering_from_source=row.get("covering_from_source", ""),
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
    rp.append(["Комментарий", "Парсятся только линейки из target_lines.csv. Материал берется из target_lines.csv при наличии."])
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

_SCAMBIO_BAD_IMAGE_HINTS = (
    "logo",
    "icon",
    "pixel",
    "placeholder",
    "loader",
    "catalog.svg",
    "search_loop",
    "colors_images",
    "sale",
    "new",
    "hit",
    "arrow",
    "cart",
    "profil",
    "korob",
    "nalichnik",
    "dobor",
    "razrez",
    "sechen",
)

_SCAMBIO_BAD_MODEL_PREFIXES = (
    "коробка", "наличник", "добор", "погонаж", "телескоп",
    "притвор", "капитель", "карниз", "розетка", "плинтус",
    "подпятник", "профиль", "комплект", "разрез", "сечение",
    "образец", "цвет", "покрытие",
)


def _normalized_model_key(model_name: str) -> str:
    cleaned, _ = normalize_dver_model_name(model_name or "")
    cleaned = re.sub(r"\.{2,}", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" \t\r\n.,;:-")
    return normalize(cleaned)


def _is_bad_image_url(image_url: str) -> bool:
    if not image_url:
        return False
    low = image_url.lower()
    return any(h in low for h in _SCAMBIO_BAD_IMAGE_HINTS) or any(k in low for k in BAD_IMAGE_KEYWORDS)


def _is_suspicious_model_name(model_name: str) -> bool:
    n = normalize(model_name)
    if not n:
        return True
    if n.startswith("дверь "):
        return False
    return any(n == p or n.startswith(p + " ") for p in _SCAMBIO_BAD_MODEL_PREFIXES)


def _row_image_score(row: DoorItem) -> int:
    score = 0
    low = (row.image_url or "").lower()
    if row.status == "found":
        score += 40
    if low:
        score += 20
    if "/xml/images/" in low:
        score += 100
    if low.endswith(".jpg") or low.endswith(".jpeg"):
        score += 15
    if _is_bad_image_url(low):
        score -= 80
    if row.source_url:
        score += 5
    return score


def dedup_by_model(rows: list[DoorItem]) -> tuple[list[DoorItem], list[dict]]:
    grouped: dict[tuple[str, str, str], list[DoorItem]] = defaultdict(list)
    for row in rows:
        key = (
            normalize(row.manufacturer),
            normalize(row.line),
            _normalized_model_key(row.model_name),
        )
        grouped[key].append(row)

    selected: list[DoorItem] = []
    duplicates: list[dict] = []
    for (prod, line, model_key), group in grouped.items():
        ranked = sorted(group, key=_row_image_score, reverse=True)
        winner = ranked[0]
        selected.append(winner)
        for dup in ranked[1:]:
            duplicates.append({
                "producer": dup.manufacturer,
                "target_line": dup.line,
                "model_name": dup.model_name,
                "selected_image_url": winner.image_url,
                "rejected_image_url": dup.image_url,
                "source_url": dup.source_url,
                "reason": "duplicate_model_same_line",
            })
    return selected, duplicates


def write_duplicates_report(path: Path, rows: list[dict]) -> None:
    fields = [
        "producer",
        "target_line",
        "model_name",
        "selected_image_url",
        "rejected_image_url",
        "source_url",
        "reason",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields, delimiter=";")
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in fields})


def write_image_audit(path: Path, rows: list[dict]) -> None:
    fields = [
        "target_line",
        "model_name",
        "selected_image_url",
        "rejected_image_urls",
        "image_status",
        "reason",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields, delimiter=";")
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in fields})


def print_scambio_console_summary(
    rows: list[DoorItem],
    targets: list[TargetLine],
    site_log: dict | None = None,
) -> None:
    target_lines = [t.target_line for t in targets]
    grouped: dict[str, list[DoorItem]] = defaultdict(list)
    for r in rows:
        grouped[normalize(r.line)].append(r)

    print("\n  СВОДКА СКАМБИО:")
    print(f"    строк всего: {len(rows)}")
    print(f"    target_line найдено: {sum(1 for t in target_lines if normalize(t) in grouped)}/{len(target_lines)}")

    print("    моделей по линейкам:")
    for t in target_lines:
        cnt = len({normalize(r.model_name) for r in grouped.get(normalize(t), [])})
        print(f"      - {t}: {cnt}")

    missing_image = sum(1 for r in rows if not r.image_url)
    needs_review = sum(1 for r in rows if r.status == "needs_review")
    print(f"    строк без фото: {missing_image}")
    print(f"    needs_review: {needs_review}")
    if site_log:
        print(f"    удалено комплектующих: {site_log.get('removed_components_count', 0)}")
        print(f"    плохих фото отклонено: {site_log.get('rejected_bad_images_count', 0)}")

    not_found_lines = [t for t in target_lines if normalize(t) not in grouped]
    print("    not_found линейки:")
    if not_found_lines:
        print("      - " + "; ".join(not_found_lines))
    else:
        print("      - нет")

    suspicious_urls = Counter(
        (r.image_url or "").strip()
        for r in rows
        if r.image_url and _is_bad_image_url(r.image_url)
    )
    suspicious_models = Counter(
        (r.model_name or "").strip()
        for r in rows
        if _is_suspicious_model_name(r.model_name)
    )

    print("    топ-20 подозрительных image_url:")
    for url, cnt in suspicious_urls.most_common(20):
        print(f"      - {cnt}x {url}")
    if not suspicious_urls:
        print("      - нет")

    print("    топ-20 подозрительных model_name:")
    for mn, cnt in suspicious_models.most_common(20):
        print(f"      - {cnt}x {mn}")
    if not suspicious_models:
        print("      - нет")

    important_lines = {
        "смальта деко",
        "смальта белла",
        "смальта галант",
        "смальта нова",
        "смальта шарм",
        "инвизибл",
    }
    print("    примеры строк (до 10) для Смальта*/Инвизибл:")
    sample_rows = [
        r for r in rows
        if normalize(r.line) in important_lines
    ][:10]
    for r in sample_rows:
        print(f"      - [{r.line}] {r.model_name} | status={r.status}")
    if not sample_rows:
        print("      - нет")

    print("    примеры image_url (10):")
    images = [r.image_url for r in rows if r.image_url][:10]
    for url in images:
        print(f"      - {url}")
    if not images:
        print("      - нет")

def write_qa_report(path: Path, all_rows: list[DoorItem], targets: list[TargetLine],
                    report: dict, site_rows: dict[str, list[DoorItem]]) -> None:
    grouped: dict[tuple[str, str], list[DoorItem]] = {}
    for row in all_rows:
        grouped.setdefault((normalize(row.manufacturer), normalize(row.line)), []).append(row)

    checks_map: dict[tuple[str, str], list[str]] = {}
    for man_name, man_info in (report.get("sites") or {}).items():
        cchecks = man_info.get("collection_checks", {})
        for line, urls in cchecks.items():
            checks_map[(normalize(man_name), normalize(line))] = list(urls or [])

    with path.open("w", encoding="utf-8-sig", newline="") as f:
        fields = [
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
        w = csv.DictWriter(f, fieldnames=fields, delimiter=";")
        w.writeheader()

        for target in targets:
            key = (normalize(target.manufacturer), normalize(target.target_line))
            rows = grouped.get(key, [])
            expected_material = target.covering_from_source or target.material_manual or ""

            rows_count = len(rows)
            unique_models = len({_normalized_model_key(r.model_name) for r in rows})
            with_image_count = sum(1 for r in rows if r.image_url)
            missing_image_count = rows_count - with_image_count
            needs_review_count = sum(1 for r in rows if r.status == "needs_review")
            bad_image_count = sum(1 for r in rows if _is_bad_image_url(r.image_url))

            if rows_count == 0:
                status = "not_found"
                checked_urls = checks_map.get(key, [])
                if normalize(target.manufacturer) == normalize("Скамбио Порте"):
                    comment = "checked_xml_feed_and_sitemap_no_models_found"
                    if checked_urls:
                        checked_joined = " | ".join(checked_urls[:4])
                        comment = f"{comment}; checked: {checked_joined}"
                elif checked_urls:
                    checked_joined = " | ".join(checked_urls[:4])
                    comment = f"collection_not_found_on_site; checked: {checked_joined}"
                else:
                    comment = "collection_not_found_on_site; no collection checks recorded"
            elif bad_image_count > 0 or missing_image_count > 0 or needs_review_count > 0:
                status = "needs_review"
                comment = (
                    f"with_image={with_image_count}; "
                    f"missing_image={missing_image_count}; "
                    f"bad_image={bad_image_count}; "
                    f"needs_review={needs_review_count}"
                )
            else:
                status = "found"
                comment = "ok"

            w.writerow({
                "producer": target.manufacturer,
                "target_line": target.target_line,
                "expected_material": expected_material,
                "rows_count": rows_count,
                "unique_models_count": unique_models,
                "with_image_count": with_image_count,
                "missing_image_count": missing_image_count,
                "needs_review_count": needs_review_count,
                "bad_image_count": bad_image_count,
                "status": status,
                "comment": comment,
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
    if args.only:
        n_only = normalize(args.only)
        targets = [t for t in targets if normalize(t.manufacturer) == n_only]
        if not targets:
            print(f"[WARN] Manufacturer '{args.only}' not found in targets.")
            sys.exit(1)
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
        print(f"\n  Filtered to manufacturer: {args.only}")

    if args.limit_sites:
        manufacturers_to_process = manufacturers_to_process[:args.limit_sites]
        print(f"  Limited to {args.limit_sites} site(s)")

    image_audit_rows: list[dict] = []

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
            image_audit_rows.extend(site_log.get("image_audit", []))
            site_log["image_audit_count"] = len(site_log.get("image_audit", []))
            site_log.pop("image_audit", None)
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

    # Final dedup by producer + target_line + normalized model_name
    all_rows, duplicates_rows = dedup_by_model(all_rows)

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
    report["duplicates_removed"] = len(duplicates_rows)
    report["image_audit_rows"] = len(image_audit_rows)
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
    write_image_audit(out_dir / "image_audit.csv", image_audit_rows)
    write_duplicates_report(out_dir / "duplicates_report.csv", duplicates_rows)
    write_xlsx(out_dir / "doors_catalog.xlsx", all_rows, targets, report)
    (out_dir / "run_report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    # Clean up partials
    for f in out_dir.glob("partial_*"):
        f.unlink()

    # For Scambio-only checks keep output folder flat (no stale subfolders like "скамбио порте")
    if args.only and normalize(args.only) == normalize("Скамбио Порте"):
        for child in out_dir.iterdir():
            if child.is_dir():
                shutil.rmtree(child, ignore_errors=True)

    # Summary with status breakdown
    found_total = sum(1 for r in all_rows if r.status == "found")
    review_total = sum(1 for r in all_rows if r.status == "needs_review")
    nf_count = report["total_targets"] - report["total_lines_found"]

    print(f"\n  DONE:")
    print(f"    Lines: {report['total_lines_found']}/{report['total_targets']} ({report['total_lines_found']/report['total_targets']*100:.1f}%)")
    print(f"    Models found: {found_total}")
    print(f"    Needs review: {review_total}")
    print(f"    Not found lines: {nf_count}")
    print(f"    Duplicates removed: {len(duplicates_rows)}")
    print(f"    Errors: {len(report['errors'])}")
    print(f"    Output: {out_dir.resolve()}")
    if args.only and normalize(args.only) == normalize("Скамбио Порте"):
        scambio_log = None
        for producer_name, site_log in report.get("sites", {}).items():
            if normalize(producer_name) == normalize("Скамбио Порте"):
                scambio_log = site_log
                break
        print_scambio_console_summary(all_rows, targets, scambio_log)
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
