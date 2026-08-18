from __future__ import annotations

import csv
import json
import re
import shutil
import struct
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parent.parent
BASE_DIR = ROOT / "output" / "final_ready_v2"
BASE_CSV = BASE_DIR / "doors_catalog.csv"
TARGETS_CSV = ROOT / "data" / "target_lines.csv"
UNMATCHED_CSV = ROOT / "output" / "olymp_model_rows_check" / "olymp_unmatched_images.csv"
OUT_DIR = ROOT / "output" / "final_ready_v3"
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

REBUILD_FIELDS = [
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
]

AUDIT_FIELDS = [
    "producer",
    "rows",
    "rows_with_image",
    "rows_without_image",
    "needs_review",
    "variant_rows",
    "generated_model_rows",
    "repeated_image_gt_10",
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


def parse_png_size(data: bytes) -> tuple[int, int] | None:
    if len(data) < 24 or data[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    return struct.unpack(">I", data[16:20])[0], struct.unpack(">I", data[20:24])[0]


def parse_gif_size(data: bytes) -> tuple[int, int] | None:
    if len(data) < 10 or not (data.startswith(b"GIF87a") or data.startswith(b"GIF89a")):
        return None
    w, h = struct.unpack("<HH", data[6:10])
    return w, h


def parse_webp_size(data: bytes) -> tuple[int, int] | None:
    if len(data) < 30 or data[:4] != b"RIFF" or data[8:12] != b"WEBP":
        return None
    kind = data[12:16]
    if kind == b"VP8 " and len(data) >= 30:
        w = struct.unpack("<H", data[26:28])[0] & 0x3FFF
        h = struct.unpack("<H", data[28:30])[0] & 0x3FFF
        return w, h
    if kind == b"VP8L" and len(data) >= 25:
        b0, b1, b2, b3 = data[21], data[22], data[23], data[24]
        w = 1 + (((b1 & 0x3F) << 8) | b0)
        h = 1 + (((b3 & 0x0F) << 10) | (b2 << 2) | ((b1 & 0xC0) >> 6))
        return w, h
    if kind == b"VP8X" and len(data) >= 30:
        w = 1 + int.from_bytes(data[24:27], "little")
        h = 1 + int.from_bytes(data[27:30], "little")
        return w, h
    return None


def parse_jpeg_size(data: bytes) -> tuple[int, int] | None:
    if len(data) < 4 or data[:2] != b"\xff\xd8":
        return None
    i = 2
    while i + 9 < len(data):
        if data[i] != 0xFF:
            i += 1
            continue
        marker = data[i + 1]
        i += 2
        if marker in (0xD8, 0xD9):
            continue
        if i + 2 > len(data):
            break
        length = struct.unpack(">H", data[i:i + 2])[0]
        if length < 2 or i + length > len(data):
            break
        if marker in {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF}:
            if i + 7 <= len(data):
                h = struct.unpack(">H", data[i + 3:i + 5])[0]
                w = struct.unpack(">H", data[i + 5:i + 7])[0]
                return w, h
        i += length
    return None


def image_size_from_bytes(data: bytes) -> tuple[int, int]:
    for parser in (parse_png_size, parse_jpeg_size, parse_webp_size, parse_gif_size):
        out = parser(data)
        if out:
            return out
    return 0, 0


def bad_image_url(url: str) -> bool:
    low = (url or "").lower()
    if not low:
        return True
    bad = (
        "logo",
        "icon",
        "badge",
        "sticker",
        "label",
        "sale",
        "promo",
        "banner",
        "search",
        "arrow",
        "placeholder",
        "no_image",
        "no-image",
        "pixel",
        "open.png",
        "catalog.svg",
        "sprite",
        "loader",
        "favicon",
        "/resize/20x/",
        "resize/20x",
        "/cover/240x180/",
        "preview",
        "close.png",
        "tl.png",
        "tg.png",
        "whatsapp",
        "telegram",
        "free-icon",
        "instagram",
        "vk-social",
    )
    return any(x in low for x in bad)


def tilda_restore(url: str) -> str:
    u = (url or "").strip()
    if not u:
        return ""
    u = re.sub(r"/-/resizeb?/\d+x/", "/", u, flags=re.I)
    u = re.sub(r"/-/cover/\d+x\d+/", "/", u, flags=re.I)
    u = re.sub(r"/-/format/webp/", "/", u, flags=re.I)
    if "optim.tildacdn.com" in u:
        u = u.replace("optim.tildacdn.com", "static.tildacdn.com")
    if "thb.tildacdn.com" in u:
        u = u.replace("thb.tildacdn.com", "static.tildacdn.com")
    return u


def cleanup_olymp_name(name: str) -> str:
    n = (name or "").strip()
    n = n.replace("_", " ")
    n = re.sub(r"\.(png|jpg|jpeg|webp)+$", "", n, flags=re.I)
    n = re.sub(r"^\d+\s+", "", n)
    n = re.sub(r"\s+", " ", n).strip(" -_")
    return n


def is_trash_name(name: str) -> bool:
    n = norm(name)
    if not n:
        return True
    bad_parts = (
        "фабрика дверей",
        "каталог",
        "купить",
        "меню",
        "главная",
        "смотреть",
        "подробнее",
        "undefined",
        "image",
        "photo",
    )
    if any(p in n for p in bad_parts):
        return True
    if re.fullmatch(r"[._\-\d\s]+", (name or "").strip()):
        return True
    return False


COMPONENT_RE = re.compile(
    r"^(добор|коробка|наличник|комплект|погонаж|телескоп|притвор\w*|капитель|карниз|розетка|плинтус|подпятник|профиль|разрез|сечение|образец|цвет|покрытие)\b",
    re.I,
)


def is_component_name(name: str) -> bool:
    return bool(COMPONENT_RE.search((name or "").strip()))


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


def fetch_image_meta(session: requests.Session, url: str, cache: dict[str, tuple[int, str, int, int]]) -> tuple[int, str, int, int]:
    if url in cache:
        return cache[url]
    try:
        r = session.get(url, timeout=(8, 25), allow_redirects=True, stream=True)
        status = r.status_code
        ctype = (r.headers.get("Content-Type") or "").lower()
        data = b""
        for chunk in r.iter_content(chunk_size=8192):
            data += chunk
            if len(data) >= 196608:
                break
        w, h = image_size_from_bytes(data) if ctype.startswith("image/") else (0, 0)
        out = (status, ctype, w, h)
    except Exception:
        out = (0, "", 0, 0)
    cache[url] = out
    return out


def main() -> None:
    if OUT_DIR.exists():
        shutil.rmtree(OUT_DIR)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    CHECK_DIR.mkdir(parents=True, exist_ok=True)

    base_rows = [{k: r.get(k, "") for k in CATALOG_FIELDS} for r in read_csv(BASE_CSV)]
    targets = read_csv(TARGETS_CSV)
    unmatched = read_csv(UNMATCHED_CSV) if UNMATCHED_CSV.exists() else []

    old_counts = Counter(r.get("producer", "") for r in base_rows if r.get("producer", ""))
    haus_before = [r for r in base_rows if r.get("producer") == "Хаусдорс"]
    olymp_before = [r for r in base_rows if r.get("producer") == "Олимп"]

    # 1) remove component rows globally
    removed_components = []
    rows_no_components = []
    for r in base_rows:
        if is_component_name(r.get("model_name", "")):
            removed_components.append(r)
            continue
        rows_no_components.append(r)

    # 2) olymp cleanup + unmatched preserve
    non_olymp = [r for r in rows_no_components if r.get("producer") != "Олимп"]
    olymp_rows = [r for r in rows_no_components if r.get("producer") == "Олимп"]
    olymp_cleaned = []
    preserved_from_trash = []
    photo_idx_by_line = Counter()
    for r in olymp_rows:
        row = dict(r)
        row["model_name"] = cleanup_olymp_name(row.get("model_name", ""))
        row["image_url"] = tilda_restore(row.get("image_url", ""))
        if bad_image_url(row.get("image_url", "")):
            continue
        if is_trash_name(row.get("model_name", "")):
            photo_idx_by_line[row.get("target_line", "")] += 1
            row["model_name"] = f"{row.get('target_line','')} — фото {photo_idx_by_line[row.get('target_line','')]}"
            row["status"] = "needs_review"
            row["comment"] = "no_visible_caption_photo_preserved_for_manual_naming"
            preserved_from_trash.append(row)
            continue
        olymp_cleaned.append(row)

    # add preserved unmatched photos
    by_line_defaults = {}
    for r in olymp_cleaned or olymp_rows:
        line = r.get("target_line", "")
        if line and line not in by_line_defaults:
            by_line_defaults[line] = r

    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36",
            "Accept-Language": "ru,en;q=0.9",
        }
    )
    meta_cache: dict[str, tuple[int, str, int, int]] = {}
    existing_olymp_images = {r.get("image_url", "") for r in olymp_cleaned if r.get("image_url", "")}
    preserve_rows = []
    preserve_counter_by_line = Counter()
    for u in unmatched:
        line = u.get("target_line", "")
        if not line:
            continue
        img = tilda_restore(u.get("image_url", ""))
        if not img or img in existing_olymp_images or bad_image_url(img):
            continue

        card_text = norm(u.get("card_text", ""))
        # reject components/menu/footer blocks
        if any(x in card_text for x in ("погонаж", "карниз", "характеристик", "фабрика межкомнатныхи входных дверей")):
            continue

        status, ctype, w, h = fetch_image_meta(session, img, meta_cache)
        if status != 200 or not ctype.startswith("image/"):
            continue
        if w < 260 or h < 420:
            continue
        ratio = w / max(h, 1)
        if ratio > 1.2 or ratio < 0.2:
            continue

        defaults = by_line_defaults.get(line)
        if not defaults:
            # try from olymp_before
            for r in olymp_before:
                if r.get("target_line", "") == line:
                    defaults = r
                    break
        if not defaults:
            continue
        preserve_counter_by_line[line] += 1
        model_name = f"{line} — фото {preserve_counter_by_line[line]}"
        row = {
            "door_type": defaults.get("door_type", "межкомнатная"),
            "producer": "Олимп",
            "target_line": line,
            "model_name": model_name,
            "description": defaults.get("description", ""),
            "material": defaults.get("material", ""),
            "coating": defaults.get("coating", ""),
            "image_url": img,
            "source_url": defaults.get("source_url", ""),
            "status": "needs_review",
            "comment": "no_visible_caption_photo_preserved_for_manual_naming",
        }
        preserve_rows.append(row)
        existing_olymp_images.add(img)

    olymp_final = olymp_cleaned + preserved_from_trash + preserve_rows

    # dedupe olymp by producer+line+name (keep found with image)
    ranked = sorted(
        olymp_final,
        key=lambda r: (
            norm(r.get("producer", "")),
            norm(r.get("target_line", "")),
            norm(r.get("model_name", "")),
            0 if (r.get("status") == "found" and r.get("image_url", "")) else 1,
            -len(r.get("image_url", "")),
        ),
    )
    dedup = {}
    for r in ranked:
        key = (r.get("producer", ""), r.get("target_line", ""), norm(r.get("model_name", "")))
        if key not in dedup:
            dedup[key] = r
    olymp_final = list(dedup.values())
    # avoid reusing identical image for many olymp rows inside one collection
    dedup_img = {}
    for r in sorted(olymp_final, key=lambda x: (norm(x.get("target_line", "")), 0 if x.get("status") == "found" else 1)):
        ikey = (r.get("target_line", ""), r.get("image_url", ""))
        if not r.get("image_url", ""):
            continue
        if ikey in dedup_img:
            continue
        dedup_img[ikey] = r
    olymp_final = list(dedup_img.values())

    final_rows = non_olymp + olymp_final
    rows_sorted = sorted(final_rows, key=lambda r: (norm(r.get("producer", "")), norm(r.get("target_line", "")), norm(r.get("model_name", ""))))

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

    # image reparse diff for Hausdoors + Olymp
    image_reparse_rows = []
    for producer in ("Хаусдорс", "Олимп"):
        old = [r for r in base_rows if r.get("producer") == producer]
        new = [r for r in rows_sorted if r.get("producer") == producer]
        old_map = {(r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", "")): r for r in old}
        new_map = {(r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", "")): r for r in new}
        keys = set(old_map) | set(new_map)
        for key in sorted(keys):
            o = old_map.get(key)
            n = new_map.get(key)
            if o and not n:
                action = "removed_old_row"
                reason = "component_removed" if is_component_name(o.get("model_name", "")) else "trash_or_generated_removed"
                row = {
                    "producer": producer,
                    "target_line": o.get("target_line", ""),
                    "model_name": o.get("model_name", ""),
                    "source_url": o.get("source_url", ""),
                    "old_image_url": o.get("image_url", ""),
                    "new_image_url": "",
                    "action": action,
                    "reason": reason,
                    "http_status": "",
                    "content_type": "",
                    "width": "",
                    "height": "",
                    "status_after": "",
                    "comment_after": "",
                }
            elif n and not o:
                reason = "no_visible_caption_photo_preserved_for_manual_naming" if "no_visible_caption_photo_preserved_for_manual_naming" in (n.get("comment") or "") else "added_after_cleanup"
                row = {
                    "producer": producer,
                    "target_line": n.get("target_line", ""),
                    "model_name": n.get("model_name", ""),
                    "source_url": n.get("source_url", ""),
                    "old_image_url": "",
                    "new_image_url": n.get("image_url", ""),
                    "action": "added_row",
                    "reason": reason,
                    "http_status": "",
                    "content_type": "",
                    "width": "",
                    "height": "",
                    "status_after": n.get("status", ""),
                    "comment_after": n.get("comment", ""),
                }
            else:
                if (
                    o.get("image_url", "") == n.get("image_url", "")
                    and o.get("status", "") == n.get("status", "")
                    and o.get("comment", "") == n.get("comment", "")
                ):
                    continue
                row = {
                    "producer": producer,
                    "target_line": n.get("target_line", ""),
                    "model_name": n.get("model_name", ""),
                    "source_url": n.get("source_url", ""),
                    "old_image_url": o.get("image_url", ""),
                    "new_image_url": n.get("image_url", ""),
                    "action": "updated_row",
                    "reason": "name_or_image_or_status_updated",
                    "http_status": "",
                    "content_type": "",
                    "width": "",
                    "height": "",
                    "status_after": n.get("status", ""),
                    "comment_after": n.get("comment", ""),
                }
            image_reparse_rows.append(row)
    write_csv(OUT_DIR / "image_reparse_report.csv", image_reparse_rows, REPARSE_FIELDS)

    # rebuild report
    haus_after = [r for r in rows_sorted if r.get("producer") == "Хаусдорс"]
    olymp_after = [r for r in rows_sorted if r.get("producer") == "Олимп"]
    rebuild_rows = [
        {
            "producer": "Хаусдорс",
            "old_rows": len(haus_before),
            "new_rows": len(haus_after),
            "old_variant_rows": sum(1 for r in haus_before if re.search(r"\bvariant\s*\d+\b|\bвариант\s*\d+\b", r.get("model_name", ""), flags=re.I)),
            "new_variant_rows": sum(1 for r in haus_after if re.search(r"\bvariant\s*\d+\b|\bвариант\s*\d+\b", r.get("model_name", ""), flags=re.I)),
            "real_model_rows": sum(1 for r in haus_after if r.get("status") == "found"),
            "generated_model_rows": sum(1 for r in haus_after if "no_visible_caption" in (r.get("comment") or "")),
            "rows_with_image": sum(1 for r in haus_after if r.get("image_url", "")),
            "rows_without_image": sum(1 for r in haus_after if not r.get("image_url", "")),
            "repeated_image_gt_10_count": 0,
            "needs_review_count": sum(1 for r in haus_after if r.get("status") == "needs_review"),
            "comment": "components_removed",
        },
        {
            "producer": "Олимп",
            "old_rows": len(olymp_before),
            "new_rows": len(olymp_after),
            "old_variant_rows": sum(1 for r in olymp_before if re.search(r"\bvariant\s*\d+\b|\bвариант\s*\d+\b|\bмодель\s*\d+\b", r.get("model_name", ""), flags=re.I)),
            "new_variant_rows": sum(1 for r in olymp_after if re.search(r"\bvariant\s*\d+\b|\bвариант\s*\d+\b|\bмодель\s*\d+\b", r.get("model_name", ""), flags=re.I)),
            "real_model_rows": sum(
                1
                for r in olymp_after
                if "no_visible_caption_photo_preserved_for_manual_naming" not in (r.get("comment") or "") and r.get("status") == "found"
            ),
            "generated_model_rows": sum(
                1 for r in olymp_after if "no_visible_caption_photo_preserved_for_manual_naming" in (r.get("comment") or "")
            ),
            "rows_with_image": sum(1 for r in olymp_after if r.get("image_url", "")),
            "rows_without_image": sum(1 for r in olymp_after if not r.get("image_url", "")),
            "repeated_image_gt_10_count": 0,
            "needs_review_count": sum(1 for r in olymp_after if r.get("status") == "needs_review"),
            "comment": "trash_removed_and_unmatched_photo_preserved",
        },
    ]
    write_csv(OUT_DIR / "rebuild_report.csv", rebuild_rows, REBUILD_FIELDS)

    # check by producer all 10
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
    audit_rows = []
    for p in producers:
        sub = [r for r in rows_sorted if r.get("producer") == p]
        stem = producer_file.get(p, norm(p).replace(" ", "_"))
        write_csv(CHECK_DIR / f"{stem}.csv", sub, CATALOG_FIELDS)
        write_xlsx(CHECK_DIR / f"{stem}.xlsx", sub, CATALOG_FIELDS, title=stem)
        img_counts = Counter(r.get("image_url", "") for r in sub if r.get("image_url", ""))
        repeated_gt10 = sum(1 for r in sub if r.get("image_url", "") and img_counts.get(r.get("image_url", ""), 0) > 10)
        variant_rows = sum(
            1 for r in sub if re.search(r"\bvariant\s*\d+\b|\bвариант\s*\d+\b|\bмодель\s*\d+\b", r.get("model_name", ""), flags=re.I)
        )
        generated_rows = sum(
            1
            for r in sub
            if "no_visible_caption_photo_preserved_for_manual_naming" in (r.get("comment") or "")
        )
        audit_rows.append(
            {
                "producer": p,
                "rows": len(sub),
                "rows_with_image": sum(1 for r in sub if r.get("image_url", "")),
                "rows_without_image": sum(1 for r in sub if not r.get("image_url", "")),
                "needs_review": sum(1 for r in sub if r.get("status") == "needs_review"),
                "variant_rows": variant_rows,
                "generated_model_rows": generated_rows,
                "repeated_image_gt_10": repeated_gt10,
                "comment": "updated" if p in {"Олимп", "Хаусдорс"} else "unchanged",
            }
        )
    write_csv(CHECK_DIR / "final_problem_sites_audit.csv", audit_rows, AUDIT_FIELDS)

    # run report + unchanged check
    base_by_prod = defaultdict(list)
    new_by_prod = defaultdict(list)
    for r in base_rows:
        base_by_prod[r.get("producer", "")].append(r)
    for r in rows_sorted:
        new_by_prod[r.get("producer", "")].append(r)
    unchanged = {}
    for p in ("Энтро", "АСД", "Скамбио Порте", "Лорд", "Аэлита", "Комфорт", "Двери регионов", "Термодур"):
        b = sorted(base_by_prod.get(p, []), key=lambda r: (r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", ""), r.get("image_url", "")))
        n = sorted(new_by_prod.get(p, []), key=lambda r: (r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", ""), r.get("image_url", "")))
        unchanged[p] = {"rows_before": len(b), "rows_after": len(n), "exact_match": b == n}

    run_report = {
        "generated_at": datetime.now().isoformat(),
        "mode": "final_ready_v3",
        "source_catalog": str(BASE_CSV),
        "old_total_rows": len(base_rows),
        "total_rows": len(rows_sorted),
        "producers_count": len(producers),
        **coverage,
        "rows_with_image": sum(1 for r in rows_sorted if r.get("image_url", "")),
        "rows_without_image": sum(1 for r in rows_sorted if not r.get("image_url", "")),
        "needs_review_count": sum(1 for r in rows_sorted if r.get("status") == "needs_review"),
        "rows_by_producer": dict(sorted(Counter(r.get("producer", "") for r in rows_sorted if r.get("producer", "")).items())),
        "hausdoors_removed_components_count": sum(1 for r in removed_components if r.get("producer") == "Хаусдорс"),
        "olymp_photo_preserved_for_manual_naming": len(preserve_rows),
        "olymp_unmatched_input_rows": len(unmatched),
        "unchanged_producers_check": unchanged,
    }
    (OUT_DIR / "run_report.json").write_text(json.dumps(run_report, ensure_ascii=False, indent=2), encoding="utf-8")

    print("Built output/final_ready_v3")
    print(
        json.dumps(
            {
                "total_rows": len(rows_sorted),
                "rows_by_producer": run_report["rows_by_producer"],
                "haus_removed_components": run_report["hausdoors_removed_components_count"],
                "olymp_preserved_photos": len(preserve_rows),
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
