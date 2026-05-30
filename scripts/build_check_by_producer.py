from __future__ import annotations

import csv
import re
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parent.parent
INPUT_CSV = ROOT / "output" / "doors_catalog.csv"
SCAMBIO_OVERRIDE_CSV = ROOT / "output" / "scambio_check" / "doors_catalog.csv"
OUT_DIR = ROOT / "output" / "check_by_producer"

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

PRODUCER_TO_FILE = {
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

SUSPICIOUS_IMAGE_HINTS = (
    "logo", "icon", "badge", "label", "sale", "new", "hit", "termo", "cart", "search", "arrow",
    "marker", "promo", "sticker", "catalog.svg", "search_loop", "colors_images", "open.png",
    "no_image", "placeholder", "blank", "sprite", "loader", "pixel", "favicon", "thumb",
    "thumbnail", "small", "mini", "preview", "resize", "96x", "100x", "150x", "200x",
)

SUSPICIOUS_MODEL_TOKENS = (
    "каталог", "поиск", "меню", "корзина", "избранное", "сравнение", "комплект", "коробка",
    "наличник", "добор", "погонаж", "телескоп", "притвор", "капитель", "карниз", "розетка",
    "плинтус", "подпятник", "профиль", "разрез", "сечение", "образец", "цвет", "покрытие",
)


def normalize(text: str) -> str:
    text = (text or "").lower().replace("ё", "е")
    text = re.sub(r"[^a-zа-я0-9]+", " ", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip()


def read_rows(path: Path) -> list[dict]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")
        rows = []
        for r in reader:
            rows.append({k: (v or "").strip() for k, v in r.items()})
    return rows


def write_csv(path: Path, rows: list[dict]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDS, delimiter=";")
        writer.writeheader()
        for row in rows:
            writer.writerow({k: row.get(k, "") for k in FIELDS})


def write_xlsx(path: Path, rows: list[dict], title: str = "catalog") -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = title[:31]
    ws.append(FIELDS)
    for row in rows:
        ws.append([row.get(k, "") for k in FIELDS])

    widths = [14, 22, 28, 42, 60, 18, 16, 60, 60, 14, 42]
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


def is_suspicious_image(url: str) -> bool:
    low = (url or "").lower()
    if not low:
        return True
    return any(h in low for h in SUSPICIOUS_IMAGE_HINTS)


def is_suspicious_model_name(name: str) -> bool:
    low = normalize(name)
    if not low:
        return True
    return any(tok in low for tok in SUSPICIOUS_MODEL_TOKENS)


def producer_key(row: dict) -> str:
    return (row.get("producer") or row.get("manufacturer") or "").strip()


def group_rows_by_producer(rows: list[dict]) -> dict[str, list[dict]]:
    grouped: dict[str, list[dict]] = defaultdict(list)
    for row in rows:
        producer = producer_key(row)
        if not producer:
            continue
        grouped[producer].append(row)
    return grouped


def make_samples_sheet(wb: Workbook, producer: str, rows: list[dict]) -> None:
    ws = wb.create_sheet(title=PRODUCER_TO_FILE.get(producer, producer)[:31])
    fields = ["sample_group"] + FIELDS
    ws.append(fields)

    first20 = rows[:20]
    with_img = [r for r in rows if r.get("image_url", "").strip()][:20]
    needs_review = [r for r in rows if (r.get("status") or "").strip() == "needs_review"][:20]
    suspicious = [r for r in rows if is_suspicious_image(r.get("image_url", "")) or is_suspicious_model_name(r.get("model_name", ""))][:20]

    buckets = [
        ("first_20", first20),
        ("with_image_20", with_img),
        ("needs_review_20", needs_review),
        ("suspicious_20", suspicious),
    ]
    for group_name, bucket_rows in buckets:
        for r in bucket_rows:
            ws.append([group_name] + [r.get(k, "") for k in FIELDS])

    widths = [18, 14, 22, 28, 42, 60, 18, 16, 60, 60, 14, 42]
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4F81BD")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def build() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    rows = read_rows(INPUT_CSV)
    if SCAMBIO_OVERRIDE_CSV.exists():
        scambio_rows = read_rows(SCAMBIO_OVERRIDE_CSV)
        if scambio_rows:
            rows = [r for r in rows if producer_key(r) != "Скамбио Порте"] + scambio_rows
    grouped = group_rows_by_producer(rows)

    summary_rows: list[dict] = []
    sample_wb = Workbook()
    sample_wb.remove(sample_wb.active)

    for producer, filename in PRODUCER_TO_FILE.items():
        producer_rows = grouped.get(producer, [])
        producer_rows = sorted(
            producer_rows,
            key=lambda r: (
                normalize(r.get("target_line", "")),
                normalize(r.get("model_name", "")),
            ),
        )

        write_csv(OUT_DIR / f"{filename}.csv", producer_rows)
        write_xlsx(OUT_DIR / f"{filename}.xlsx", producer_rows, title=filename)
        make_samples_sheet(sample_wb, producer, producer_rows)

        missing_image_count = sum(1 for r in producer_rows if not r.get("image_url", "").strip())
        needs_review_count = sum(1 for r in producer_rows if (r.get("status") or "").strip() == "needs_review")
        suspicious_image_count = sum(1 for r in producer_rows if is_suspicious_image(r.get("image_url", "")))
        suspicious_model_count = sum(1 for r in producer_rows if is_suspicious_model_name(r.get("model_name", "")))

        dup_counter = Counter(
            (normalize(r.get("target_line", "")), normalize(r.get("model_name", "")))
            for r in producer_rows
            if r.get("model_name", "").strip()
        )
        duplicate_model_name_count = sum(cnt - 1 for cnt in dup_counter.values() if cnt > 1)

        summary_rows.append(
            {
                "producer": producer,
                "rows_count": len(producer_rows),
                "target_lines_count": len({normalize(r.get("target_line", "")) for r in producer_rows if r.get("target_line", "").strip()}),
                "missing_image_count": missing_image_count,
                "needs_review_count": needs_review_count,
                "suspicious_image_count": suspicious_image_count,
                "suspicious_model_name_count": suspicious_model_count,
                "duplicate_model_name_count": duplicate_model_name_count,
                "comment": "ok" if producer_rows else "no_rows_in_current_catalog",
            }
        )

    audit_fields = [
        "producer",
        "rows_count",
        "target_lines_count",
        "missing_image_count",
        "needs_review_count",
        "suspicious_image_count",
        "suspicious_model_name_count",
        "duplicate_model_name_count",
        "comment",
    ]
    with (OUT_DIR / "audit_summary.csv").open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=audit_fields, delimiter=";")
        writer.writeheader()
        for row in summary_rows:
            writer.writerow(row)

    sample_wb.save(OUT_DIR / "samples_for_visual_check.xlsx")
    print(f"check_by_producer ready: {OUT_DIR}")


if __name__ == "__main__":
    build()
