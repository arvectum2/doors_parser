from __future__ import annotations

import csv
import json
import re
import shutil
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parent.parent
SRC_CATALOG = ROOT / "output" / "doors_catalog.csv"
TARGETS_CSV = ROOT / "data" / "target_lines.csv"
OUT_DIR = ROOT / "output" / "scambio_check"


BAD_MODEL_PREFIXES = (
    "комплект",
    "коробка",
    "наличник",
    "добор",
    "погонаж",
    "телескоп",
    "притвор",
    "капитель",
    "карниз",
    "розетка",
    "плинтус",
    "подпятник",
    "профиль",
    "разрез",
    "сечение",
)

BAD_IMAGE_HINTS = (
    "colors_images",
    "catalog.svg",
    "search_loop",
    "pixel.gif",
    "loader",
    "placeholder",
    "no_image",
    "logo",
    "icon",
    "arrow",
    "cart",
    "nalichnik",
    "korob",
    "dobor",
    "profil",
    "razrez",
    "sechen",
)

FIELDS = [
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


def normalize(text: str) -> str:
    text = (text or "").lower().replace("ё", "е")
    text = re.sub(r"[^a-zа-я0-9]+", " ", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip()


def clean_model_name(name: str) -> str:
    t = name or ""
    t = re.sub(r"\s*[–—-]?\s*распродажа\s*[.,;:!?]*", " ", t, flags=re.IGNORECASE)
    t = t.replace("×", "x").replace("Х", "x").replace("х", "x").replace("*", "x")
    t = re.sub(r"\bразмер\s*[:\-]?\s*\d[\d.]*\s*x\s*\d[\d.]*(\s*x\s*\d[\d.]*)?\b", " ", t, flags=re.I)
    t = re.sub(r"\b\d[\d.]*\s*x\s*\d[\d.]*(\s*x\s*\d[\d.]*)?\b", " ", t, flags=re.I)
    t = re.sub(r"\.{2,}", " ", t)
    t = re.sub(r"\s+", " ", t).strip(" \t\r\n.,;:-")
    return t


def non_door_model(name: str) -> bool:
    n = normalize(name)
    if not n:
        return True
    if n.startswith("дверь "):
        return False
    return any(n == p or n.startswith(p + " ") for p in BAD_MODEL_PREFIXES)


def bad_image(url: str) -> bool:
    if not url:
        return False
    low = url.lower()
    return any(k in low for k in BAD_IMAGE_HINTS)


def extract_code(text: str, prefix: str) -> str | None:
    pattern = re.compile(rf"\b{prefix}\s*[-–—]?\s*([0-9]+(?:\.[0-9]+)?)\b", re.I)
    m = pattern.search(text)
    return m.group(1) if m else None


def extract_smalta_num(text: str) -> str | None:
    # explicit "Смальта-Лайн 07", "Смальта 23.2"
    m = re.search(
        r"смальта\s*[-–—]?\s*(?:деко|б[еэ]лла|галант|лайн|нова|риф|шарм)?\s*[-–—]?\s*([0-9]+(?:\.[0-9]+)?)\b",
        text,
        re.I,
    )
    return m.group(1) if m else None


def infer_line(model: str, desc: str, source_url: str) -> tuple[str | None, bool]:
    text = f"{model} {desc} {source_url}".replace("ё", "е")
    low = text.lower()
    low = low.replace("xром", "хром")

    # Strong direct matches first.
    strong_map = [
        (r"смальта\s*[-–—]?\s*деко", "Смальта Деко"),
        (r"смальта\s*[-–—]?\s*б[еэ]лла", "Смальта Белла"),
        (r"смальта\s*[-–—]?\s*галант", "Смальта Галант"),
        (r"смальта\s*[-–—]?\s*лайн", "Смальта Лайн"),
        (r"смальта\s*[-–—]?\s*нова", "Смальта Нова"),
        (r"смальта\s*[-–—]?\s*риф", "Смальта Риф"),
        (r"смальта\s*[-–—]?\s*шарм", "Смальта Шарм"),
        (r"\bпсл\b", "ПСЛ"),
        (r"\bпск\b", "ПСК"),
        (r"\bпсу\b", "ПСУ"),
        (r"\bпст\b", "ПСТ"),
        (r"\bпсе\b", "ПСЕ"),
        (r"\bпде\b", "ПДЕ"),
        (r"\bинвиз|invisible\b", "Инвизибл"),
    ]
    for pat, line in strong_map:
        if re.search(pat, low, re.I):
            return line, True

    if re.search(r"абс\s*кромк", low, re.I):
        return "П АБС кромка Черный", True

    if "хром" in low and ("молдинг" in low or "кромка" in low):
        return "ПХ Алл кромка Хром", False

    if re.search(r"\bпх\b|\bпx\b", low, re.I):
        if "хром" in low:
            return "ПХ Алл кромка Хром", True
        if "черн" in low:
            return "ПХ Алл кромка Черный", True
        # keep uncertain split explicit
        return "ПХ Алл кромка Черный", False

    psc_num = extract_code(text, "ПСЦ")
    if psc_num:
        try:
            number = float(psc_num)
        except ValueError:
            number = 0.0
        # conservative: low numbers mostly classic, high mostly modern
        return ("ПСЦ Классика" if number < 20 else "ПСЦ Модерн"), False

    ps_num = extract_code(text, "ПС")
    if ps_num:
        if ps_num in {"5", "14"}:
            return "ПСЛ", False
        if "сан-ремо" in low:
            return "ПСЛ", False
        return "ПС", True

    p_num = extract_code(text, "П")
    if p_num:
        # П-21/22 with ABS should not fall into plain П
        if re.search(r"\bп\s*[-–—]?\s*(21|22)\b", low) and "абс" in low:
            return "П АБС кромка Черный", False
        return "П", True

    # Generic Smalta rows without explicit collection name.
    sm_num = extract_smalta_num(text)
    if sm_num:
        if sm_num in {"07", "08", "7", "8"}:
            return "Смальта Лайн", True
        if sm_num in {"201", "202", "203", "204", "208.1", "208.2", "209.2", "210"}:
            return "Смальта Риф", True
        if sm_num in {"1", "01", "2", "02"}:
            # backed by collection snippets for Смальта-Шарм
            return "Смальта Шарм", False
        if sm_num in {"12"}:
            return "Смальта Белла", False
        if sm_num in {"104"}:
            return "Смальта Галант", False
        if sm_num in {"14"}:
            return "Смальта Нова", False
        # default for numbered smalta models
        return "Смальта Деко", False

    return None, False


def read_targets() -> list[dict]:
    with TARGETS_CSV.open("r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f, delimiter=";"))
    return [r for r in rows if (r.get("manufacturer") or "").strip() == "Скамбио Порте"]


def write_csv(path: Path, rows: list[dict], fields: list[str]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields, delimiter=";")
        w.writeheader()
        for row in rows:
            w.writerow({k: row.get(k, "") for k in fields})


def write_xlsx(path: Path, rows: list[dict]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "parsed_catalog"
    headers = [
        "Тип двери",
        "Производитель",
        "Линейка",
        "Название",
        "Описание",
        "Материал",
        "Покрытие",
        "Ссылка на фото",
        "Ссылка на карточку",
        "Статус",
        "Комментарий",
    ]
    ws.append(headers)
    for row in rows:
        ws.append(
            [
                row.get("door_type", ""),
                row.get("producer", ""),
                row.get("target_line", ""),
                row.get("model_name", ""),
                row.get("description", ""),
                row.get("material", ""),
                row.get("coating", ""),
                row.get("image_url", ""),
                row.get("source_url", ""),
                row.get("status", ""),
                row.get("comment", ""),
            ]
        )
    widths = [16, 22, 28, 40, 70, 18, 16, 60, 60, 14, 40]
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


def rebuild() -> None:
    if OUT_DIR.exists():
        for child in OUT_DIR.iterdir():
            if child.is_dir():
                shutil.rmtree(child)
            else:
                child.unlink(missing_ok=True)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    targets = read_targets()
    target_lines = [t["target_line"] for t in targets]
    material_map = {t["target_line"]: (t.get("covering_from_source") or t.get("material_manual") or "").strip() for t in targets}

    source_rows: list[dict] = []
    with SRC_CATALOG.open("r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f, delimiter=";"):
            if (row.get("producer") or "").strip() != "Скамбио Порте":
                continue
            source_rows.append(row)

    # Fallback source for Invisible when product pages are rate-limited:
    # sitemap contains real product URLs with article IDs and xml/images.
    try:
        sm_text = requests.get("https://dver.com/xml/sitemap.xml", timeout=30).text
        sm_urls = re.findall(r"<loc>(.*?)</loc>", sm_text)
        for url in sm_urls:
            low = url.lower()
            if "invisible" not in low:
                continue
            art = ""
            m_art = re.search(r"art-(F[0-9]+)", url, re.I)
            if m_art:
                art = m_art.group(1).upper()
            model_name = "Дверь Инвизибл"
            if art:
                model_name = f"Дверь Инвизибл {art}"
            image_url = f"https://dver.com/xml/images/{art}.jpeg" if art else ""
            source_rows.append(
                {
                    "door_type": "межкомнатная",
                    "producer": "Скамбио Порте",
                    "target_line": "Инвизибл",
                    "model_name": model_name,
                    "description": "Модель из sitemap dver.com (Invisible)",
                    "material": "Под покраску",
                    "coating": "",
                    "image_url": image_url,
                    "source_url": url,
                    "status": "found" if image_url else "needs_review",
                    "comment": "source=sitemap_fallback",
                }
            )
    except Exception:
        pass

    result_rows: list[dict] = []
    image_audit: list[dict] = []
    duplicates: list[dict] = []

    grouped: dict[tuple[str, str], list[dict]] = defaultdict(list)

    for row in source_rows:
        model_raw = (row.get("model_name") or "").strip()
        desc = (row.get("description") or "").strip()
        src = (row.get("source_url") or "").strip()
        img = (row.get("image_url") or "").strip()
        comment = (row.get("comment") or "").strip()

        if "collection_not_found_on_site" in comment.lower():
            # ignore synthetic placeholders from previous runs
            continue

        model_name = clean_model_name(model_raw)
        if non_door_model(model_name):
            continue

        inferred_line, strong = infer_line(model_name, desc, src)
        if inferred_line is None:
            continue
        if inferred_line not in material_map:
            continue

        # image selection from available value (single-source rebuild)
        rejected = []
        status = "found"
        reason = "selected_product_image"
        selected_image = img
        if not img:
            status = "needs_review"
            reason = "no_good_image_found"
        elif bad_image(img):
            rejected.append(img)
            selected_image = ""
            status = "needs_review"
            reason = "rejected_component_image"

        audit_reason = reason
        if not strong and status == "found":
            audit_reason = "selected_product_image"
            comment = f"{comment}; heuristic_line_assignment".strip("; ").strip()
        elif not strong and status == "needs_review":
            comment = f"{comment}; heuristic_line_assignment; {reason}".strip("; ").strip()
        elif status == "needs_review":
            comment = f"{comment}; {reason}".strip("; ").strip()

        rebuilt = {
            "door_type": (row.get("door_type") or "межкомнатная").strip() or "межкомнатная",
            "producer": "Скамбио Порте",
            "target_line": inferred_line,
            "model_name": model_name,
            "description": desc,
            "material": material_map.get(inferred_line, ""),
            "coating": (row.get("coating") or "").strip(),
            "image_url": selected_image,
            "source_url": src,
            "status": status,
            "comment": comment,
        }

        dedup_key = (
            normalize(rebuilt["producer"]),
            normalize(rebuilt["target_line"]),
            normalize(rebuilt["model_name"]),
        )
        grouped[dedup_key].append(rebuilt)
        image_audit.append(
            {
                "target_line": inferred_line,
                "model_name": model_name,
                "old_image_url": img,
                "selected_image_url": selected_image,
                "rejected_image_urls": " | ".join(rejected),
                "image_status": status,
                "reason": audit_reason if selected_image else reason,
            }
        )

    for key, items in grouped.items():
        # prefer row with image
        ranked = sorted(items, key=lambda r: (1 if r["image_url"] else 0, 1 if r["status"] == "found" else 0), reverse=True)
        winner = ranked[0]
        result_rows.append(winner)
        for dup in ranked[1:]:
            duplicates.append(
                {
                    "producer": dup["producer"],
                    "target_line": dup["target_line"],
                    "model_name": dup["model_name"],
                    "selected_image_url": winner["image_url"],
                    "rejected_image_url": dup["image_url"],
                    "source_url": dup["source_url"],
                    "reason": "duplicate_model_same_line",
                }
            )

    # keep only target lines from spec
    result_rows = [r for r in result_rows if r["target_line"] in target_lines]

    # QA report for all 21 lines
    by_line: dict[str, list[dict]] = defaultdict(list)
    for row in result_rows:
        by_line[row["target_line"]].append(row)

    qa_rows: list[dict] = []
    for line in target_lines:
        rows = by_line.get(line, [])
        with_image = sum(1 for r in rows if r["image_url"])
        missing = len(rows) - with_image
        needs_review = sum(1 for r in rows if r["status"] == "needs_review")
        bad_img = sum(1 for r in rows if bad_image(r["image_url"]))
        if len(rows) == 0:
            status = "not_found"
            comment = "no matched products in source catalog + dver.com/xml/sitemap.xml snapshot"
        elif needs_review > 0 or missing > 0 or bad_img > 0:
            status = "needs_review"
            comment = f"with_image={with_image}; missing_image={missing}; needs_review={needs_review}; bad_image={bad_img}"
        else:
            status = "found"
            comment = "ok"
        qa_rows.append(
            {
                "producer": "Скамбио Порте",
                "target_line": line,
                "expected_material": material_map.get(line, ""),
                "rows_count": len(rows),
                "unique_models_count": len({normalize(r['model_name']) for r in rows}),
                "with_image_count": with_image,
                "missing_image_count": missing,
                "needs_review_count": needs_review,
                "bad_image_count": bad_img,
                "status": status,
                "comment": comment,
            }
        )

    # files
    write_csv(OUT_DIR / "doors_catalog.csv", result_rows, FIELDS)
    write_xlsx(OUT_DIR / "doors_catalog.xlsx", result_rows)
    write_csv(
        OUT_DIR / "qa_report.csv",
        qa_rows,
        [
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
        ],
    )
    write_csv(
        OUT_DIR / "image_audit.csv",
        image_audit,
        [
            "target_line",
            "model_name",
            "old_image_url",
            "selected_image_url",
            "rejected_image_urls",
            "image_status",
            "reason",
        ],
    )
    write_csv(
        OUT_DIR / "duplicates_report.csv",
        duplicates,
        [
            "producer",
            "target_line",
            "model_name",
            "selected_image_url",
            "rejected_image_url",
            "source_url",
            "reason",
        ],
    )

    report = {
        "generated_at": datetime.now().isoformat(),
        "mode": "scambio_rebuild_from_existing_catalog",
        "source_catalog": str(SRC_CATALOG),
        "total_rows": len(result_rows),
        "target_lines_total": len(target_lines),
        "target_lines_found": sum(1 for line in target_lines if len(by_line.get(line, [])) > 0),
        "missing_image_rows": sum(1 for r in result_rows if not r["image_url"]),
        "needs_review_rows": sum(1 for r in result_rows if r["status"] == "needs_review"),
        "duplicates_removed": len(duplicates),
    }
    (OUT_DIR / "run_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    # console stats required before any global merge/archive
    print("SCAMBIO SUMMARY")
    print(f"rows_total={len(result_rows)}")
    for line in target_lines:
        print(f"{line}: {len(by_line.get(line, []))}")
    print(f"rows_without_image={report['missing_image_rows']}")
    print(f"needs_review={report['needs_review_rows']}")
    print("image_examples_top10:")
    examples = [r["image_url"] for r in result_rows if r["image_url"]][:10]
    for ex in examples:
        print(ex)
    print("suspicious_model_names_top20:")
    suspicious_models = [
        r["model_name"]
        for r in result_rows
        if any(x in normalize(r["model_name"]) for x in ("...", "комплект", "коробка", "наличник", "добор", "профиль"))
    ][:20]
    for item in suspicious_models:
        print(item)
    print("suspicious_image_urls_top20:")
    suspicious_images = [r["image_url"] for r in result_rows if bad_image(r["image_url"])][:20]
    for item in suspicious_images:
        print(item)


if __name__ == "__main__":
    rebuild()
