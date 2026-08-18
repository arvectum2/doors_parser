from __future__ import annotations

import csv
import json
import re
import shutil
import struct
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parent.parent
IN_CATALOG = ROOT / "output" / "final_images_fix" / "doors_catalog.csv"
OUT_DIR = ROOT / "output" / "final_images_fix"
TARGETS_CSV = ROOT / "data" / "target_lines.csv"

TARGET_PRODUCERS = {"Хаусдорс", "Энтро", "Олимп", "АСД"}
STRICT_FIX_PRODUCERS = {"Хаусдорс", "Энтро", "Олимп"}

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

PHOTO_FIX_SUMMARY_FIELDS = [
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


def normalize_model_key(text: str) -> str:
    t = normalize(text)
    t = re.sub(r"\b(дверь|door|model|модель)\b", " ", t)
    return re.sub(r"\s+", " ", t).strip()


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


def parse_png_size(data: bytes) -> tuple[int, int] | None:
    if len(data) < 24 or data[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    w = struct.unpack(">I", data[16:20])[0]
    h = struct.unpack(">I", data[20:24])[0]
    return w, h


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
    if len(data) < 4 or data[0:2] != b"\xff\xd8":
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


class ImageMeta:
    def __init__(self, url: str, status: int, ctype: str, width: int, height: int, is_image: bool, is_html: bool, error: str = ""):
        self.url = url
        self.status = status
        self.ctype = ctype
        self.width = width
        self.height = height
        self.is_image = is_image
        self.is_html = is_html
        self.error = error


def fetch_image_meta(session: requests.Session, url: str, cache: dict[str, ImageMeta]) -> ImageMeta:
    if not url:
        return ImageMeta("", 0, "", 0, 0, False, False, "empty_url")
    if url in cache:
        return cache[url]
    try:
        r = session.get(url, timeout=(8, 30), allow_redirects=True, stream=True)
        ctype = (r.headers.get("Content-Type") or "").lower()
        data = b""
        for chunk in r.iter_content(chunk_size=8192):
            data += chunk
            if len(data) >= 131072:
                break
        is_image = ctype.startswith("image/")
        is_html = "text/html" in ctype or (data[:512].lower().find(b"<html") != -1)
        w, h = (0, 0)
        if is_image and data:
            w, h = image_size_from_bytes(data)
        meta = ImageMeta(url, r.status_code, ctype, w, h, is_image, is_html, "")
    except Exception as exc:
        meta = ImageMeta(url, 0, "", 0, 0, False, False, str(exc))
    cache[url] = meta
    return meta


def tilda_variants(url: str) -> list[str]:
    out = [url]
    if "tildacdn.com" in url:
        u = re.sub(r"/-/resizeb?/\d+x/", "/", url, flags=re.I)
        out.append(u)
        if "thb.tildacdn.com" in u:
            out.append(u.replace("thb.tildacdn.com", "static.tildacdn.com"))
        if "thb.tildacdn.com" in url:
            out.append(url.replace("thb.tildacdn.com", "static.tildacdn.com"))
    uniq = []
    for u in out:
        if u not in uniq:
            uniq.append(u)
    return uniq


def extract_variant_number(model_name: str) -> int | None:
    m = re.search(r"вариант\s*(\d+)$", normalize(model_name))
    if m:
        return int(m.group(1))
    m2 = re.search(r"variant\s*(\d+)$", model_name.lower())
    if m2:
        return int(m2.group(1))
    return None


def bad_hausdoors_url(url: str) -> bool:
    low = (url or "").lower()
    if not low:
        return True
    if "/empty/" in low:
        return True
    if "noroot" in low:
        return True
    if low.endswith(".svg"):
        return True
    if any(x in low for x in ("logo", "icon", "placeholder", "pixel", "sprite", "loader", "search", "arrow", "cart", "open.png")):
        return True
    if any(x in low for x in ("ral", "ncs", "sample", "swatch", "pattern", "palette", "покры", "цвет")):
        return True
    base = low.rsplit("/", 1)[-1]
    if base in {"tl.png", "tg.png", "vk.png", "wa.png"}:
        return True
    if re.match(r"^[a-z]{1,3}\.(png|jpg|jpeg|webp)$", base):
        return True
    return False


def bad_olymp_url(url: str) -> bool:
    low = (url or "").lower()
    if not low:
        return True
    if "promo" in low:
        return True
    if any(x in low for x in ("logo", "icon", "placeholder", "pixel", "sprite", "loader", "search", "arrow", "cart", "open.png")):
        return True
    return False


def html_fetch(session: requests.Session, url: str, cache: dict[str, tuple[bool, str]]) -> tuple[bool, str]:
    if not url:
        return False, ""
    if url in cache:
        return cache[url]
    try:
        r = session.get(url, timeout=(8, 30), allow_redirects=True)
        if r.status_code >= 400:
            out = (False, "")
        else:
            out = (True, r.text)
    except Exception:
        out = (False, "")
    cache[url] = out
    return out


def extract_page_image_candidates(page_url: str, html: str) -> list[str]:
    soup = BeautifulSoup(html, "lxml")
    urls: list[str] = []
    seen: set[str] = set()

    def add(raw: str) -> None:
        if not raw:
            return
        raw = raw.strip().strip("'\"")
        if not raw:
            return
        if raw.startswith("//"):
            raw = "https:" + raw
        full = urljoin(page_url, raw)
        if full not in seen:
            seen.add(full)
            urls.append(full)

    for tag in soup.find_all("meta", property="og:image"):
        add(tag.get("content", ""))
    for img in soup.find_all("img"):
        add(img.get("data-original", ""))
        add(img.get("data-src", ""))
        add(img.get("src", ""))
    for raw in re.findall(r"background-image\s*:\s*url\(([^)]+)\)", html, flags=re.I):
        add(raw)
    for raw in re.findall(r"(https?://[^\"'\s>]+)", html):
        if any(ext in raw.lower() for ext in (".jpg", ".jpeg", ".png", ".webp")):
            add(raw)
    for raw in re.findall(r"(//[^\"'\s>]+)", html):
        if "tildacdn.com" in raw and any(ext in raw.lower() for ext in (".jpg", ".jpeg", ".png", ".webp")):
            add(raw)

    out: list[str] = []
    for u in urls:
        low = u.lower()
        if not any(ext in low for ext in (".jpg", ".jpeg", ".png", ".webp")):
            continue
        out.append(u)
    return out


def hausdoors_candidates_for_url(
    session: requests.Session,
    source_url: str,
    page_cache: dict[str, tuple[bool, str]],
    meta_cache: dict[str, ImageMeta],
) -> list[tuple[str, ImageMeta]]:
    ok, html = html_fetch(session, source_url, page_cache)
    if not ok:
        return []
    raw = extract_page_image_candidates(source_url, html)
    cand: list[tuple[str, ImageMeta, float]] = []
    seen = set()
    for u in raw:
        for v in tilda_variants(u):
            if v in seen:
                continue
            seen.add(v)
            if bad_hausdoors_url(v):
                continue
            meta = fetch_image_meta(session, v, meta_cache)
            if meta.status != 200 or not meta.is_image or meta.is_html:
                continue
            if meta.width < 360 or meta.height < 360:
                continue
            score = 0.0
            low = v.lower()
            if "static.tildacdn.com" in low:
                score += 20
            if "/-/resize/" in low or "thb.tildacdn.com" in low:
                score -= 20
            if any(x in low for x in ("invisible", "inv_", "complanar", "neo", "line", "classic", "mirror", "black", "white")):
                score += 15
            area = meta.width * meta.height
            if area > 700000:
                score += 15
            elif area > 300000:
                score += 8
            if abs(meta.width - meta.height) < 40:
                score += 3
            cand.append((v, meta, score))
    cand.sort(key=lambda x: x[2], reverse=True)
    return [(u, m) for u, m, _ in cand]


def recalc_qa(rows: list[dict], target_rows: list[dict]) -> tuple[list[dict], list[dict], dict]:
    grouped: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for row in rows:
        grouped[(normalize(row.get("producer", "")), normalize(row.get("target_line", "")))].append(row)

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
        key = (normalize(producer), normalize(line))
        sub = grouped.get(key, [])
        if not sub and normalize(producer) == normalize("АСД") and normalize(line) == normalize("Все модели"):
            sub = grouped.get((normalize("АСД"), normalize("Уточнить коллекцию")), [])

        with_image = sum(1 for r in sub if r.get("image_url", "").strip())
        missing = len(sub) - with_image
        needs_review = sum(1 for r in sub if (r.get("status") or "").strip() == "needs_review")
        bad_img = sum(1 for r in sub if "/empty/" in (r.get("image_url") or "").lower())
        unique_models = len({normalize_model_key(r.get("model_name", "")) for r in sub if r.get("model_name", "").strip()})

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
            if missing > 0 or needs_review > 0 or bad_img > 0:
                status = "needs_review"
                comment = f"with_image={with_image}; missing_image={missing}; needs_review={needs_review}; bad_image={bad_img}"
            else:
                status = "found"
                comment = "ok"
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


def main() -> None:
    rows = [{k: (r.get(k, "") or "").strip() for k in CATALOG_FIELDS} for r in read_csv(IN_CATALOG)]
    before = [{k: r.get(k, "") for k in CATALOG_FIELDS} for r in rows]

    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
            "Accept-Language": "ru,en;q=0.9",
        }
    )
    page_cache: dict[str, tuple[bool, str]] = {}
    meta_cache: dict[str, ImageMeta] = {}

    reparse_rows: list[dict] = []
    haus_problem_rows: list[dict] = []
    entro_problem_rows: list[dict] = []
    olymp_problem_rows: list[dict] = []

    before_with_image = Counter()
    after_with_image = Counter()
    replaced = Counter()
    removed = Counter()
    not_found = Counter()
    suspicious_left = Counter()

    for r in rows:
        p = r.get("producer", "")
        if p in TARGET_PRODUCERS and r.get("image_url", "").strip():
            before_with_image[p] += 1

    # 1) Hausdoors: rebuild from source pages, ban /empty/, avoid one-image fallback storms.
    haus_groups: dict[str, list[dict]] = defaultdict(list)
    for r in rows:
        if r.get("producer") == "Хаусдорс":
            haus_groups[r.get("source_url", "")].append(r)

    for src, grp in haus_groups.items():
        cands = hausdoors_candidates_for_url(session, src, page_cache, meta_cache)
        # Cap per image so one fallback is not stretched to hundreds of models.
        max_per_image = 12
        capacity: dict[str, int] = {u: max_per_image for u, _m in cands}
        ordered = sorted(grp, key=lambda x: (extract_variant_number(x.get("model_name", "")) or 10_000, normalize(x.get("model_name", ""))))
        for idx, row in enumerate(ordered):
            old_image = row.get("image_url", "").strip()
            old_status = row.get("status", "")
            old_comment = row.get("comment", "")
            new_image = ""
            reason = "no_model_image_found"
            action = "image_not_found"
            selected_meta = ImageMeta("", 0, "", 0, 0, False, False, "not_found")

            if cands:
                chosen = None
                vnum = extract_variant_number(row.get("model_name", ""))
                start_idx = (vnum - 1) if vnum else idx
                for shift in range(len(cands)):
                    u, m = cands[(start_idx + shift) % len(cands)]
                    if capacity.get(u, 0) > 0:
                        chosen = (u, m)
                        break
                if chosen:
                    new_image, selected_meta = chosen
                    capacity[new_image] -= 1
                    action = "replaced_image" if new_image != old_image else "kept_old_image"
                    reason = "collection_image_used"
                else:
                    reason = "no_model_image_found"

            if not new_image:
                row["image_url"] = ""
                row["status"] = "needs_review"
                row["comment"] = append_comment(old_comment, "no_model_image_found")
                if old_image:
                    removed["Хаусдорс"] += 1
                    action = "removed_bad_image"
                not_found["Хаусдорс"] += 1
                haus_problem_rows.append({**row, "reason": reason, "old_image_url": old_image})
            else:
                row["image_url"] = new_image
                row["status"] = "found" if old_status != "needs_review" else old_status
                row["comment"] = append_comment(old_comment, "collection_image_used")
                if new_image != old_image:
                    replaced["Хаусдорс"] += 1

            reparse_rows.append(
                {
                    "producer": "Хаусдорс",
                    "target_line": row.get("target_line", ""),
                    "model_name": row.get("model_name", ""),
                    "source_url": src,
                    "old_image_url": old_image,
                    "new_image_url": row.get("image_url", ""),
                    "action": action,
                    "reason": reason,
                    "http_status": selected_meta.status,
                    "content_type": selected_meta.ctype,
                    "width": selected_meta.width,
                    "height": selected_meta.height,
                    "status_after": row.get("status", ""),
                    "comment_after": row.get("comment", ""),
                }
            )

    # 2) Entro: keep cache/thumb only as collection image, mark clearly for review.
    for row in rows:
        if row.get("producer") != "Энтро":
            continue
        old_image = row.get("image_url", "").strip()
        old_comment = row.get("comment", "")
        action = "kept_old_image"
        reason = "full_door_image_ok"
        meta = ImageMeta("", 0, "", 0, 0, False, False, "")

        if "/wp-content/cache/thumb/" in old_image.lower():
            meta = fetch_image_meta(session, old_image, meta_cache)
            if meta.status == 200 and meta.is_image and not meta.is_html and meta.width >= 400 and meta.height >= 400:
                row["comment"] = append_comment(old_comment, "collection_image_used")
                row["comment"] = append_comment(row["comment"], "full_door_image_ok")
            else:
                row["status"] = "needs_review"
                row["comment"] = append_comment(old_comment, "cache_thumb_unverified")
                row["comment"] = append_comment(row["comment"], "collection_image_used")
                reason = "cache_thumb_unverified"
                entro_problem_rows.append({**row, "reason": reason, "old_image_url": old_image})
        else:
            if not old_image:
                row["status"] = "needs_review"
                row["comment"] = append_comment(old_comment, "no_model_image_found")
                not_found["Энтро"] += 1
                action = "image_not_found"
                reason = "no_model_image_found"
                entro_problem_rows.append({**row, "reason": reason, "old_image_url": old_image})

        reparse_rows.append(
            {
                "producer": "Энтро",
                "target_line": row.get("target_line", ""),
                "model_name": row.get("model_name", ""),
                "source_url": row.get("source_url", ""),
                "old_image_url": old_image,
                "new_image_url": row.get("image_url", ""),
                "action": action,
                "reason": reason,
                "http_status": meta.status,
                "content_type": meta.ctype,
                "width": meta.width,
                "height": meta.height,
                "status_after": row.get("status", ""),
                "comment_after": row.get("comment", ""),
            }
        )

    # 3) Olymp: promo URLs either replace from source or clear.
    olymp_groups: dict[str, list[dict]] = defaultdict(list)
    for r in rows:
        if r.get("producer") == "Олимп":
            olymp_groups[r.get("source_url", "")].append(r)

    for src, grp in olymp_groups.items():
        ok, html = html_fetch(session, src, page_cache)
        source_cands: list[tuple[str, ImageMeta]] = []
        if ok:
            raw = extract_page_image_candidates(src, html)
            seen = set()
            for u in raw:
                for v in tilda_variants(u):
                    if v in seen:
                        continue
                    seen.add(v)
                    if bad_olymp_url(v):
                        continue
                    m = fetch_image_meta(session, v, meta_cache)
                    if m.status == 200 and m.is_image and not m.is_html and m.width >= 350 and m.height >= 350:
                        source_cands.append((v, m))

        for row in grp:
            old_image = row.get("image_url", "").strip()
            old_comment = row.get("comment", "")
            action = "kept_old_image"
            reason = "full_door_image_ok"
            meta = ImageMeta("", 0, "", 0, 0, False, False, "")
            if not old_image:
                row["status"] = "needs_review"
                row["comment"] = append_comment(old_comment, "no_model_image_found")
                not_found["Олимп"] += 1
                action = "image_not_found"
                reason = "no_model_image_found"
                olymp_problem_rows.append({**row, "reason": reason, "old_image_url": old_image})
            elif "promo" in old_image.lower():
                replacement = source_cands[0][0] if source_cands else ""
                if replacement and replacement != old_image:
                    row["image_url"] = replacement
                    row["comment"] = append_comment(old_comment, "image_reparsed")
                    replaced["Олимп"] += 1
                    action = "replaced_image"
                    reason = "better_product_image_found"
                    meta = source_cands[0][1]
                else:
                    row["image_url"] = ""
                    row["status"] = "needs_review"
                    row["comment"] = append_comment(old_comment, "no_model_image_found")
                    removed["Олимп"] += 1
                    not_found["Олимп"] += 1
                    action = "removed_bad_image"
                    reason = "promo_unverified_removed"
                    olymp_problem_rows.append({**row, "reason": reason, "old_image_url": old_image})
            reparse_rows.append(
                {
                    "producer": "Олимп",
                    "target_line": row.get("target_line", ""),
                    "model_name": row.get("model_name", ""),
                    "source_url": src,
                    "old_image_url": old_image,
                    "new_image_url": row.get("image_url", ""),
                    "action": action,
                    "reason": reason,
                    "http_status": meta.status,
                    "content_type": meta.ctype,
                    "width": meta.width,
                    "height": meta.height,
                    "status_after": row.get("status", ""),
                    "comment_after": row.get("comment", ""),
                }
            )

    # 4) ASD: only ensure empty image rows are needs_review.
    for row in rows:
        if row.get("producer") != "АСД":
            continue
        if not row.get("image_url", "").strip():
            row["status"] = "needs_review"
            row["comment"] = append_comment(row.get("comment", ""), "no_model_image_found")
            not_found["АСД"] += 1

    # 5) Hausdoors URL normalization: prefer static/original over thb/resize.
    for row in rows:
        if row.get("producer") != "Хаусдорс":
            continue
        img = row.get("image_url", "").strip()
        if not img:
            continue
        best = img
        best_score = -10_000
        best_meta = None
        for cand in tilda_variants(img):
            m = fetch_image_meta(session, cand, meta_cache)
            if m.status != 200 or not m.is_image or m.is_html:
                continue
            s = 0
            low = cand.lower()
            if "static.tildacdn.com" in low:
                s += 10
            if "thb.tildacdn.com" in low:
                s -= 10
            if "/-/resize/" in low:
                s -= 15
            s += min((m.width * m.height) / 100000, 20)
            if s > best_score:
                best_score = s
                best = cand
                best_meta = m
        if best != img:
            row["image_url"] = best
            row["comment"] = append_comment(row.get("comment", ""), "image_reparsed")
            replaced["Хаусдорс"] += 1
            reparse_rows.append(
                {
                    "producer": "Хаусдорс",
                    "target_line": row.get("target_line", ""),
                    "model_name": row.get("model_name", ""),
                    "source_url": row.get("source_url", ""),
                    "old_image_url": img,
                    "new_image_url": best,
                    "action": "replaced_image",
                    "reason": "better_product_image_found",
                    "http_status": best_meta.status if best_meta else 0,
                    "content_type": best_meta.ctype if best_meta else "",
                    "width": best_meta.width if best_meta else 0,
                    "height": best_meta.height if best_meta else 0,
                    "status_after": row.get("status", ""),
                    "comment_after": row.get("comment", ""),
                }
            )

    # Post-pass suspicious counters for strict producers.
    for r in rows:
        p = r.get("producer", "")
        if p not in STRICT_FIX_PRODUCERS:
            continue
        img = r.get("image_url", "").strip().lower()
        if not img:
            continue
        if p == "Хаусдорс" and "/empty/" in img:
            suspicious_left[p] += 1
        if p == "Энтро" and "/wp-content/cache/thumb/" in img and "full_door_image_ok" not in r.get("comment", ""):
            suspicious_left[p] += 1
        if p == "Олимп" and "promo" in img and "false_positive_ok" not in r.get("comment", ""):
            suspicious_left[p] += 1

    # Sort, export.
    rows_sorted = sorted(
        rows,
        key=lambda r: (normalize(r.get("producer", "")), normalize(r.get("target_line", "")), normalize(r.get("model_name", ""))),
    )

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

    target_rows = read_csv(TARGETS_CSV)
    qa_rows, not_found_rows, coverage = recalc_qa(rows_sorted, target_rows)
    write_csv(OUT_DIR / "qa_report.csv", qa_rows, QA_FIELDS)
    write_csv(OUT_DIR / "not_found.csv", not_found_rows, NOT_FOUND_FIELDS)

    # Keep existing duplicates report unchanged.
    dup_src = OUT_DIR / "duplicates_report.csv"
    if not dup_src.exists():
        write_csv(
            dup_src,
            [],
            [
                "producer",
                "target_line",
                "normalized_model_name",
                "duplicates_count_before",
                "kept_model_name",
                "kept_image_url",
                "removed_image_urls",
                "reason",
            ],
        )

    write_csv(OUT_DIR / "image_reparse_report.csv", reparse_rows, REPARSE_FIELDS)

    # Problem files.
    check_dir = OUT_DIR / "check_by_producer"
    check_dir.mkdir(parents=True, exist_ok=True)
    prob_fields = CATALOG_FIELDS + ["reason", "old_image_url"]
    write_csv(check_dir / "hausdoors_problem_images.csv", haus_problem_rows, prob_fields)
    write_csv(check_dir / "entro_problem_images.csv", entro_problem_rows, prob_fields)
    write_csv(check_dir / "olymp_problem_images.csv", olymp_problem_rows, prob_fields)

    # Producer slices for quick review.
    producer_to_file = {"Олимп": "olymp", "АСД": "asd", "Энтро": "entro", "Хаусдорс": "hausdoors"}
    by_prod: dict[str, list[dict]] = defaultdict(list)
    for r in rows_sorted:
        if r.get("producer", "") in producer_to_file:
            by_prod[r["producer"]].append(r)
    summary_rows = []
    for producer in ("Олимп", "АСД", "Энтро", "Хаусдорс"):
        sub = by_prod.get(producer, [])
        write_csv(check_dir / f"{producer_to_file[producer]}.csv", sub, CATALOG_FIELDS)
        write_xlsx(check_dir / f"{producer_to_file[producer]}.xlsx", sub, CATALOG_FIELDS, title=producer_to_file[producer])
        rows_after = sum(1 for r in sub if r.get("image_url", "").strip())
        summary_rows.append(
            {
                "producer": producer,
                "rows_total": len(sub),
                "rows_with_image_before": before_with_image.get(producer, 0),
                "rows_with_image_after": rows_after,
                "replaced_images_count": replaced.get(producer, 0),
                "removed_bad_images_count": removed.get(producer, 0),
                "image_not_found_count": sum(1 for r in sub if not r.get("image_url", "").strip()),
                "suspicious_images_remaining_count": suspicious_left.get(producer, 0),
                "comment": "needs_manual_visual_check" if suspicious_left.get(producer, 0) else "ok",
            }
        )
        after_with_image[producer] = rows_after
    write_csv(check_dir / "photo_fix_audit_summary.csv", summary_rows, PHOTO_FIX_SUMMARY_FIELDS)

    # image_audit from diff.
    before_map = {(i,): b for i, b in enumerate(before)}
    image_audit_rows = []
    for idx, row in enumerate(rows):
        old = before_map[(idx,)]
        if old.get("image_url", "") != row.get("image_url", "") or old.get("status", "") != row.get("status", "") or old.get("comment", "") != row.get("comment", ""):
            image_audit_rows.append(
                {
                    "producer": row.get("producer", ""),
                    "target_line": row.get("target_line", ""),
                    "model_name": row.get("model_name", ""),
                    "old_image_url": old.get("image_url", ""),
                    "new_image_url": row.get("image_url", ""),
                    "source_url": row.get("source_url", ""),
                    "image_status_before": old.get("status", ""),
                    "image_status_after": row.get("status", ""),
                    "action": "changed",
                    "comment": row.get("comment", ""),
                }
            )
    write_csv(
        OUT_DIR / "image_audit.csv",
        image_audit_rows,
        ["producer", "target_line", "model_name", "old_image_url", "new_image_url", "source_url", "image_status_before", "image_status_after", "action", "comment"],
    )

    # Run report.
    producers_total = len({r.get("producer", "") for r in rows_sorted if r.get("producer", "").strip()})
    rows_with_image = sum(1 for r in rows_sorted if r.get("image_url", "").strip())
    rows_without_image = len(rows_sorted) - rows_with_image
    needs_review = sum(1 for r in rows_sorted if (r.get("status") or "").strip() == "needs_review")
    material_filled = sum(1 for r in rows_sorted if r.get("material", "").strip())

    run_report = {
        "generated_at": datetime.now().isoformat(),
        "mode": "photo_quality_fix_v3",
        "total_rows": len(rows_sorted),
        "producers_count": producers_total,
        **coverage,
        "rows_with_image": rows_with_image,
        "rows_without_image": rows_without_image,
        "material_filled_count": material_filled,
        "needs_review_count": needs_review,
        "photo_fix_summary": {row["producer"]: row for row in summary_rows},
        "hausdoors_empty_urls_remaining": sum(1 for r in rows_sorted if r.get("producer") == "Хаусдорс" and "/empty/" in r.get("image_url", "").lower()),
        "entro_cache_thumb_count": sum(1 for r in rows_sorted if r.get("producer") == "Энтро" and "/wp-content/cache/thumb/" in r.get("image_url", "").lower()),
        "olymp_promo_count": sum(1 for r in rows_sorted if r.get("producer") == "Олимп" and "promo" in r.get("image_url", "").lower()),
        "entro_photo_sources_checked": sorted({r.get("source_url", "") for r in rows_sorted if r.get("producer") == "Энтро"}),
        "hausdoors_suspicious_checked": len(haus_problem_rows),
        "asd_suspicious_checked": 0,
    }
    (OUT_DIR / "run_report.json").write_text(json.dumps(run_report, ensure_ascii=False, indent=2), encoding="utf-8")

    print("photo quality fix v3 completed")
    for s in summary_rows:
        print(
            f"{s['producer']}: rows={s['rows_total']} with_image_after={s['rows_with_image_after']} "
            f"replaced={s['replaced_images_count']} removed={s['removed_bad_images_count']} "
            f"empty={s['image_not_found_count']} suspicious={s['suspicious_images_remaining_count']}"
        )


if __name__ == "__main__":
    main()
