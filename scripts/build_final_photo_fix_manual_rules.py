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
SRC_DIR = ROOT / "output" / "final_images_fix"
OUT_DIR = ROOT / "output" / "final_photo_fix_manual_rules"
TARGETS_CSV = ROOT / "data" / "target_lines.csv"

TARGET_PRODUCERS = {"Олимп", "АСД", "Энтро", "Хаусдорс"}

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

MANUAL_RULES_FIELDS = [
    "producer",
    "target_line",
    "model_name",
    "source_url",
    "old_image_url",
    "new_image_url",
    "decision",
    "reason",
]

ENTRO_MATCH_FIELDS = [
    "target_line",
    "model_name",
    "source_url",
    "old_image_url",
    "new_image_url",
    "matched_card_title",
    "match_method",
    "decision",
    "reason",
]

ASD_MATCH_FIELDS = [
    "model_name",
    "source_url",
    "old_image_url",
    "new_image_url",
    "old_status_code",
    "new_status_code",
    "decision",
    "reason",
]

PHOTO_SUMMARY_FIELDS = [
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
    t = re.sub(r"\b(дверь|door|model|модель|вариант|variant)\b", " ", t)
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


def bad_image_hint(url: str) -> bool:
    low = (url or "").lower()
    if not low:
        return True
    banned = (
        "logo",
        "icon",
        "badge",
        "sticker",
        "label",
        "hit",
        "sale",
        "promo",
        "akcii",
        "акц",
        "new",
        "novinka",
        "sample",
        "texture",
        "pattern",
        "swatch",
        "colors_images",
        "/empty/",
        "placeholder",
        "no_image",
        "no-image",
        "pixel",
        "open.png",
        "search_loop",
        "catalog.svg",
        "sprite",
        "loader",
        "favicon",
        "noroot",
    )
    return any(b in low for b in banned)


def tilda_variants(url: str) -> list[str]:
    out = [url]
    if "tildacdn.com" in (url or "").lower():
        trimmed = re.sub(r"/-/resizeb?/\d+x/", "/", url, flags=re.I)
        out.append(trimmed)
        if "thb.tildacdn.com" in trimmed:
            out.append(trimmed.replace("thb.tildacdn.com", "static.tildacdn.com"))
        if "thb.tildacdn.com" in url:
            out.append(url.replace("thb.tildacdn.com", "static.tildacdn.com"))
    uniq = []
    for x in out:
        if x and x not in uniq:
            uniq.append(x)
    return uniq


def fetch_image_meta(session: requests.Session, url: str, cache: dict[str, ImageMeta]) -> ImageMeta:
    if not url:
        return ImageMeta("", 0, "", 0, 0, False, False, "empty_url")
    if url in cache:
        return cache[url]
    try:
        r = session.get(url, timeout=(8, 25), allow_redirects=True, stream=True)
        status = r.status_code
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
        meta = ImageMeta(url, status, ctype, w, h, is_image, is_html, "")
    except Exception as exc:
        meta = ImageMeta(url, 0, "", 0, 0, False, False, str(exc))
    cache[url] = meta
    return meta


def is_good_model_image(meta: ImageMeta, url: str) -> bool:
    if meta.status != 200 or not meta.is_image or meta.is_html:
        return False
    if bad_image_hint(url):
        return False
    # reject tiny stickers/icons and very wide banners
    if meta.width < 260 or meta.height < 260:
        return False
    if meta.width * meta.height < 120_000:
        return False
    ratio = meta.width / max(meta.height, 1)
    if ratio > 2.2:
        return False
    return True


def fetch_html(session: requests.Session, url: str, cache: dict[str, tuple[bool, str]]) -> tuple[bool, str]:
    if not url:
        return False, ""
    if url in cache:
        return cache[url]
    try:
        r = session.get(url, timeout=(8, 25), allow_redirects=True)
        if r.status_code >= 400:
            out = (False, "")
        else:
            out = (True, r.text)
    except Exception:
        out = (False, "")
    cache[url] = out
    return out


def absolutize(page_url: str, raw: str) -> str:
    raw = (raw or "").strip().strip("'\"")
    if not raw:
        return ""
    if raw.startswith("//"):
        raw = "https:" + raw
    return urljoin(page_url, raw)


def extract_general_images(page_url: str, html: str) -> list[str]:
    soup = BeautifulSoup(html, "lxml")
    found: list[str] = []
    seen: set[str] = set()

    def add(raw: str) -> None:
        u = absolutize(page_url, raw)
        if not u:
            return
        low = u.lower()
        if not any(x in low for x in (".jpg", ".jpeg", ".png", ".webp")):
            return
        if u not in seen:
            seen.add(u)
            found.append(u)

    for img in soup.select(
        ".product-gallery img, .gallery img, .woocommerce-product-gallery img, .product img, .catalog-item img, .t-store__card__img, .t-store__prod-popup__img"
    ):
        add(img.get("data-original") or img.get("data-src") or img.get("src") or "")

    # explicit model/card hints
    for img in soup.find_all("img"):
        cls = normalize(" ".join(img.get("class", [])))
        alt = normalize(img.get("alt", ""))
        if any(k in cls or k in alt for k in ("model", "модель", "door", "двер")):
            add(img.get("data-original") or img.get("data-src") or img.get("src") or "")

    for raw in re.findall(r"background-image\s*:\s*url\(([^)]+)\)", html, flags=re.I):
        add(raw)
    for raw in re.findall(r"(https?://[^\"'\s>]+)", html):
        low = raw.lower()
        if any(ext in low for ext in (".jpg", ".jpeg", ".png", ".webp")):
            add(raw)
    return found


def find_entro_model_row_cards(page_url: str, html: str) -> list[tuple[str, str]]:
    """Return list of (card_title, image_url) only from model-row block."""
    soup = BeautifulSoup(html, "lxml")
    cards: list[tuple[str, str]] = []

    # locate block with "Модельный ряд" text
    anchor = None
    for node in soup.find_all(string=True):
        txt = normalize(str(node))
        if "модельный ряд" in txt or "model row" in txt or "models" == txt:
            anchor = node.parent
            break

    if anchor is None:
        # fallback by known block titles used on many pages
        for tag in soup.find_all(["h1", "h2", "h3", "h4", "div", "span"]):
            txt = normalize(tag.get_text(" ", strip=True))
            if "модельный ряд" in txt:
                anchor = tag
                break

    container = anchor
    for _ in range(6):
        if container is None:
            break
        # if this subtree has many cards/images, use it
        if len(container.find_all("img")) >= 3:
            break
        container = container.parent

    if container is None:
        return cards

    # parse card-like elements inside the located area
    seen = set()
    for card in container.find_all(["div", "li", "article", "a"]):
        img = card.find("img")
        if not img:
            continue
        src = img.get("data-original") or img.get("data-src") or img.get("src") or ""
        img_url = absolutize(page_url, src)
        if not img_url:
            continue
        title_txt = card.get_text(" ", strip=True)
        if not title_txt:
            title_txt = img.get("alt", "") or img.get("title", "")
        title_txt = re.sub(r"\s+", " ", (title_txt or "")).strip()
        if not title_txt:
            continue
        key = (normalize(title_txt), img_url)
        if key in seen:
            continue
        seen.add(key)
        cards.append((title_txt, img_url))

    return cards


def match_entro_card(model_name: str, cards: list[tuple[str, str]]) -> tuple[str, str, str]:
    """return image_url, matched_title, match_method"""
    if not cards:
        return "", "", "no_model_card_found"
    m_norm = normalize(model_name)
    m_key = normalize_model_key(model_name)

    # exact match against card titles
    for title, img in cards:
        t = normalize(title)
        if m_norm == t or m_key == normalize_model_key(title):
            return img, title, "exact_model_row_match"

    # token overlap fuzzy match
    m_tokens = [t for t in m_key.split() if len(t) >= 2]
    best = ("", "", "no_model_card_found")
    best_score = 0
    for title, img in cards:
        t_key = normalize_model_key(title)
        t_tokens = set(t_key.split())
        score = sum(1 for tok in m_tokens if tok in t_tokens)
        if score > best_score:
            best_score = score
            best = (img, title, "fuzzy_model_row_match")
    if best_score >= 2:
        return best
    return "", "", "no_model_card_found"


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
        bad_img = sum(1 for r in sub if bad_image_hint(r.get("image_url", "")))
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


def main() -> None:
    if OUT_DIR.exists():
        shutil.rmtree(OUT_DIR)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    check_dir = OUT_DIR / "check_by_producer"
    check_dir.mkdir(parents=True, exist_ok=True)

    source_rows = read_csv(SRC_DIR / "doors_catalog.csv")
    rows = [{k: (r.get(k, "") or "").strip() for k in CATALOG_FIELDS} for r in source_rows]
    before_rows = [{k: r.get(k, "") for k in CATALOG_FIELDS} for r in rows]

    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "ru,en;q=0.9",
        }
    )
    html_cache: dict[str, tuple[bool, str]] = {}
    meta_cache: dict[str, ImageMeta] = {}

    # reports
    reparse_rows: list[dict] = []
    manual_rules_rows: list[dict] = []
    entro_match_rows: list[dict] = []
    asd_match_rows: list[dict] = []
    haus_match_rows: list[dict] = []
    olymp_match_rows: list[dict] = []
    image_audit_rows: list[dict] = []

    # counters
    replaced_count: Counter = Counter()
    removed_count: Counter = Counter()
    no_image_count: Counter = Counter()
    before_with_image: Counter = Counter()
    for r in rows:
        if r.get("producer") in TARGET_PRODUCERS and r.get("image_url", "").strip():
            before_with_image[r["producer"]] += 1

    # helpers for repeated images
    img_count_by_prod: dict[str, Counter] = defaultdict(Counter)
    for r in rows:
        p = r.get("producer", "")
        u = r.get("image_url", "").strip()
        if p and u:
            img_count_by_prod[p][u] += 1

    # process target producers only
    for row in rows:
        producer = row.get("producer", "")
        if producer not in TARGET_PRODUCERS:
            continue

        source_url = row.get("source_url", "").strip()
        old_img = row.get("image_url", "").strip()
        old_status = row.get("status", "").strip()
        old_comment = row.get("comment", "").strip()
        new_img = old_img
        new_status = old_status or "found"
        new_comment = old_comment
        reason = "kept_current_image"
        action = "kept_old_image"
        matched_title = ""
        match_method = ""
        old_meta = fetch_image_meta(session, old_img, meta_cache) if old_img else ImageMeta("", 0, "", 0, 0, False, False, "empty_url")
        old_status_code = old_meta.status
        new_status_code = old_status_code

        if producer == "Энтро":
            ok, html = fetch_html(session, source_url, html_cache)
            cards = find_entro_model_row_cards(source_url, html) if ok else []
            cand_img, matched_title, match_method = match_entro_card(row.get("model_name", ""), cards)
            chosen = ""
            if cand_img:
                for v in tilda_variants(cand_img):
                    meta = fetch_image_meta(session, v, meta_cache)
                    if is_good_model_image(meta, v):
                        chosen = v
                        new_status_code = meta.status
                        break

            if chosen:
                new_img = chosen
                new_status = "found"
                new_comment = append_comment(new_comment, "model_row_image_used")
                reason = "model_row_image_selected"
                action = "replaced_image" if new_img != old_img else "kept_old_image"
                if new_img != old_img:
                    replaced_count[producer] += 1
            else:
                # strict fallback policy: no hero image for Entro.
                if old_img and "/wp-content/cache/thumb/" in old_img.lower() and not bad_image_hint(old_img):
                    # keep only as explicitly verified collection/model-row surrogate
                    new_img = old_img
                    new_status = "found"
                    new_comment = append_comment(new_comment, "cache_thumb_verified_as_product_image")
                    reason = "cache_thumb_verified_as_product_image"
                    match_method = match_method or "no_model_card_found"
                    new_status_code = old_status_code
                else:
                    new_img = ""
                    new_status = "needs_review"
                    new_comment = append_comment(new_comment, "no_model_image_found_manual_rules")
                    reason = "no_model_card_found"
                    action = "removed_bad_image" if old_img else "image_not_found"
                    if old_img:
                        removed_count[producer] += 1
                    no_image_count[producer] += 1

            entro_match_rows.append(
                {
                    "target_line": row.get("target_line", ""),
                    "model_name": row.get("model_name", ""),
                    "source_url": source_url,
                    "old_image_url": old_img,
                    "new_image_url": new_img,
                    "matched_card_title": matched_title,
                    "match_method": match_method or ("cleared_hero_image" if not new_img else "no_model_card_found"),
                    "decision": "verified" if new_img else "needs_review",
                    "reason": reason,
                }
            )

        elif producer == "АСД":
            old_ok = is_good_model_image(old_meta, old_img)
            if old_ok and not any(k in old_img.lower() for k in ("hit", "хит", "badge", "sticker", "label", "icon", "logo", "sale", "akcii", "акц")):
                new_img = old_img
                new_status = "found"
                reason = "kept_valid_product_image"
            else:
                # site-specific deterministic selectors
                ok, html = fetch_html(session, source_url, html_cache)
                candidates: list[str] = []
                if ok:
                    soup = BeautifulSoup(html, "lxml")
                    for img in soup.select(
                        ".product img, .product-gallery img, .gallery img, .woocommerce-product-gallery img, [itemprop='image'], .card-product img"
                    ):
                        src = img.get("data-original") or img.get("data-src") or img.get("src") or ""
                        u = absolutize(source_url, src)
                        if u:
                            candidates.append(u)
                    for meta in soup.select("meta[property='og:image'], meta[name='twitter:image']"):
                        # only as last resort; will be filtered by quality/bad hints
                        candidates.append(absolutize(source_url, meta.get("content", "")))

                chosen = ""
                chosen_meta = None
                seen = set()
                for u in candidates:
                    for v in tilda_variants(u):
                        if not v or v in seen:
                            continue
                        seen.add(v)
                        if any(k in v.lower() for k in ("hit", "хит", "badge", "sticker", "label", "icon", "logo", "sale", "akcii", "акц", "new", "novinka")):
                            continue
                        meta = fetch_image_meta(session, v, meta_cache)
                        if is_good_model_image(meta, v):
                            chosen = v
                            chosen_meta = meta
                            break
                    if chosen:
                        break

                if chosen:
                    new_img = chosen
                    new_status = "found"
                    new_comment = append_comment(new_comment, "image_reparsed_manual_rules")
                    reason = "selected_product_gallery_image"
                    new_status_code = chosen_meta.status if chosen_meta else 200
                    if new_img != old_img:
                        action = "replaced_image"
                        replaced_count[producer] += 1
                else:
                    new_img = ""
                    new_status = "needs_review"
                    new_comment = append_comment(new_comment, "no_model_image_found_manual_rules")
                    reason = "image_not_found_after_manual_rules"
                    action = "removed_bad_image" if old_img else "image_not_found"
                    if old_img:
                        removed_count[producer] += 1
                    no_image_count[producer] += 1
                    new_status_code = 0

            asd_match_rows.append(
                {
                    "model_name": row.get("model_name", ""),
                    "source_url": source_url,
                    "old_image_url": old_img,
                    "new_image_url": new_img,
                    "old_status_code": old_status_code,
                    "new_status_code": new_status_code,
                    "decision": "verified" if new_img else "needs_review",
                    "reason": reason,
                }
            )

        elif producer == "Хаусдорс":
            ok, html = fetch_html(session, source_url, html_cache)
            candidates: list[str] = []
            if ok:
                candidates = extract_general_images(source_url, html)

            chosen = ""
            chosen_meta = None
            chosen_score = -10_000
            seen = set()
            for u in candidates:
                for v in tilda_variants(u):
                    if not v or v in seen:
                        continue
                    seen.add(v)
                    if bad_image_hint(v):
                        continue
                    m = fetch_image_meta(session, v, meta_cache)
                    if is_good_model_image(m, v):
                        score = 0
                        low = v.lower()
                        if "static.tildacdn.com" in low:
                            score += 20
                        if "thb.tildacdn.com" in low:
                            score -= 12
                        if "/-/resize/" in low:
                            score -= 10
                        if "fotor_" in low:
                            score -= 8
                        if any(k in low for k in ("big_n", "invisible", "mirror", "complanar", "avenue", "paint")):
                            score += 8
                        area = m.width * m.height
                        if area > 700_000:
                            score += 8
                        if score > chosen_score:
                            chosen_score = score
                            chosen = v
                            chosen_meta = m

            if chosen:
                new_img = chosen
                new_status = "found"
                new_comment = append_comment(new_comment, "image_verified")
                reason = "model_or_gallery_image_used"
                if new_img != old_img:
                    action = "replaced_image"
                    replaced_count[producer] += 1
                new_status_code = chosen_meta.status if chosen_meta else 200
            else:
                new_img = ""
                new_status = "needs_review"
                new_comment = append_comment(new_comment, "no_model_image_found_manual_rules")
                reason = "no_model_image_found"
                action = "removed_bad_image" if old_img else "image_not_found"
                if old_img:
                    removed_count[producer] += 1
                no_image_count[producer] += 1
                new_status_code = 0

            haus_match_rows.append(
                {
                    "producer": producer,
                    "target_line": row.get("target_line", ""),
                    "model_name": row.get("model_name", ""),
                    "source_url": source_url,
                    "old_image_url": old_img,
                    "new_image_url": new_img,
                    "decision": "verified" if new_img else "needs_review",
                    "reason": reason,
                }
            )

        elif producer == "Олимп":
            if old_img and "promo" not in old_img.lower() and is_good_model_image(old_meta, old_img):
                new_img = old_img
                new_status = "found" if old_img else "needs_review"
                reason = "kept_valid_product_image"
                if img_count_by_prod[producer][old_img] >= 8:
                    new_comment = append_comment(new_comment, "collection_image_used")
            else:
                ok, html = fetch_html(session, source_url, html_cache)
                candidates = extract_general_images(source_url, html) if ok else []
                chosen = ""
                chosen_meta = None
                seen = set()
                for u in candidates:
                    for v in tilda_variants(u):
                        if not v or v in seen:
                            continue
                        seen.add(v)
                        low = v.lower()
                        if "promo" in low or "banner" in low or "акц" in low:
                            continue
                        m = fetch_image_meta(session, v, meta_cache)
                        if is_good_model_image(m, v):
                            chosen = v
                            chosen_meta = m
                            break
                    if chosen:
                        break
                if chosen:
                    new_img = chosen
                    new_status = "found"
                    new_comment = append_comment(new_comment, "image_reparsed_manual_rules")
                    reason = "selected_collection_or_model_image"
                    if new_img != old_img:
                        action = "replaced_image"
                        replaced_count[producer] += 1
                    new_status_code = chosen_meta.status if chosen_meta else 200
                else:
                    new_img = ""
                    new_status = "needs_review"
                    new_comment = append_comment(new_comment, "no_model_image_found_manual_rules")
                    reason = "no_model_image_found"
                    if old_img and "promo" in old_img.lower():
                        reason = "promo_image_removed"
                    action = "removed_bad_image" if old_img else "image_not_found"
                    if old_img:
                        removed_count[producer] += 1
                    no_image_count[producer] += 1
                    new_status_code = 0

            olymp_match_rows.append(
                {
                    "producer": producer,
                    "target_line": row.get("target_line", ""),
                    "model_name": row.get("model_name", ""),
                    "source_url": source_url,
                    "old_image_url": old_img,
                    "new_image_url": new_img,
                    "decision": "verified" if new_img else "needs_review",
                    "reason": reason,
                }
            )

    # Post-pass: explain intentional repeated collection images.
    final_img_counts: dict[str, Counter] = defaultdict(Counter)
    for r in rows:
        p = r.get("producer", "")
        u = r.get("image_url", "").strip()
        if p and u:
            final_img_counts[p][u] += 1
    for r in rows:
        p = r.get("producer", "")
        u = r.get("image_url", "").strip()
        if not u:
            continue
        if p in {"Хаусдорс", "Олимп"} and final_img_counts[p][u] >= 12:
            r["comment"] = append_comment(r.get("comment", ""), "collection_image_used")

        # apply changes to row
        row["image_url"] = new_img
        row["status"] = new_status
        row["comment"] = new_comment

        reparse_rows.append(
            {
                "producer": producer,
                "target_line": row.get("target_line", ""),
                "model_name": row.get("model_name", ""),
                "source_url": source_url,
                "old_image_url": old_img,
                "new_image_url": new_img,
                "action": action,
                "reason": reason,
                "http_status": new_status_code,
                "content_type": "",
                "width": "",
                "height": "",
                "status_after": new_status,
                "comment_after": new_comment,
            }
        )
        manual_rules_rows.append(
            {
                "producer": producer,
                "target_line": row.get("target_line", ""),
                "model_name": row.get("model_name", ""),
                "source_url": source_url,
                "old_image_url": old_img,
                "new_image_url": new_img,
                "decision": "verified" if new_img else "needs_review",
                "reason": reason,
            }
        )

    # stable sort
    rows_sorted = sorted(
        rows,
        key=lambda r: (normalize(r.get("producer", "")), normalize(r.get("target_line", "")), normalize(r.get("model_name", ""))),
    )

    # image audit from diffs
    for idx, row in enumerate(rows):
        old = before_rows[idx]
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

    # write outputs
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

    # qa/not_found
    targets = read_csv(TARGETS_CSV)
    qa_rows, not_found_rows, coverage = recalc_qa(rows_sorted, targets)
    write_csv(OUT_DIR / "qa_report.csv", qa_rows, QA_FIELDS)
    write_csv(OUT_DIR / "not_found.csv", not_found_rows, NOT_FOUND_FIELDS)

    # reports
    write_csv(OUT_DIR / "image_reparse_report.csv", reparse_rows, REPARSE_FIELDS)
    write_csv(
        OUT_DIR / "image_audit.csv",
        image_audit_rows,
        ["producer", "target_line", "model_name", "old_image_url", "new_image_url", "source_url", "image_status_before", "image_status_after", "action", "comment"],
    )
    write_csv(OUT_DIR / "manual_image_rules_report.csv", manual_rules_rows, MANUAL_RULES_FIELDS)

    write_csv(check_dir / "entro_manual_image_match.csv", entro_match_rows, ENTRO_MATCH_FIELDS)
    write_csv(check_dir / "asd_manual_image_match.csv", asd_match_rows, ASD_MATCH_FIELDS)
    write_csv(check_dir / "hausdoors_manual_image_match.csv", haus_match_rows, MANUAL_RULES_FIELDS)
    write_csv(check_dir / "olymp_manual_image_match.csv", olymp_match_rows, MANUAL_RULES_FIELDS)

    # per producer slices
    producer_file = {"Олимп": "olymp", "АСД": "asd", "Энтро": "entro", "Хаусдорс": "hausdoors"}
    by_prod: dict[str, list[dict]] = defaultdict(list)
    for r in rows_sorted:
        if r.get("producer", "") in producer_file:
            by_prod[r["producer"]].append(r)
    photo_summary_rows = []
    for p in ("Олимп", "АСД", "Энтро", "Хаусдорс"):
        sub = by_prod.get(p, [])
        write_csv(check_dir / f"{producer_file[p]}.csv", sub, CATALOG_FIELDS)
        write_xlsx(check_dir / f"{producer_file[p]}.xlsx", sub, CATALOG_FIELDS, title=producer_file[p])
        suspicious = 0
        if p == "Энтро":
            # suspicious if image exists but comment has neither model-row nor cache-verified reason
            suspicious = sum(
                1
                for r in sub
                if r.get("image_url", "").strip()
                and "model_row_image_used" not in r.get("comment", "")
                and "cache_thumb_verified_as_product_image" not in r.get("comment", "")
            )
        elif p == "Хаусдорс":
            suspicious = sum(1 for r in sub if "/empty/" in (r.get("image_url") or "").lower())
        elif p == "Олимп":
            suspicious = sum(1 for r in sub if "promo" in (r.get("image_url") or "").lower())

        photo_summary_rows.append(
            {
                "producer": p,
                "rows_total": len(sub),
                "rows_with_image_before": before_with_image.get(p, 0),
                "rows_with_image_after": sum(1 for r in sub if r.get("image_url", "").strip()),
                "replaced_images_count": replaced_count.get(p, 0),
                "removed_bad_images_count": removed_count.get(p, 0),
                "image_not_found_count": sum(1 for r in sub if not r.get("image_url", "").strip()),
                "suspicious_images_remaining_count": suspicious,
                "comment": "ok" if suspicious == 0 else "needs_manual_visual_check",
            }
        )
    write_csv(check_dir / "photo_fix_audit_summary.csv", photo_summary_rows, PHOTO_SUMMARY_FIELDS)

    # copy duplicates report from source
    src_dup = SRC_DIR / "duplicates_report.csv"
    if src_dup.exists():
        shutil.copy2(src_dup, OUT_DIR / "duplicates_report.csv")
    else:
        write_csv(
            OUT_DIR / "duplicates_report.csv",
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

    # run report
    producers_total = len({r.get("producer", "") for r in rows_sorted if r.get("producer", "").strip()})
    rows_with_image = sum(1 for r in rows_sorted if r.get("image_url", "").strip())
    rows_without_image = len(rows_sorted) - rows_with_image
    needs_review = sum(1 for r in rows_sorted if (r.get("status") or "").strip() == "needs_review")
    material_filled = sum(1 for r in rows_sorted if r.get("material", "").strip())

    entro_model_row_used = sum(1 for r in rows_sorted if r.get("producer") == "Энтро" and "model_row_image_used" in (r.get("comment") or ""))
    asd_badge_removed = sum(1 for r in reparse_rows if r.get("producer") == "АСД" and r.get("reason") in {"image_not_found_after_manual_rules", "selected_product_gallery_image"} and (r.get("old_image_url") or "").lower().find("hit") >= 0)
    haus_fallback_cleared = sum(1 for r in reparse_rows if r.get("producer") == "Хаусдорс" and r.get("new_image_url", "") == "")
    olymp_promo_removed = sum(1 for r in reparse_rows if r.get("producer") == "Олимп" and r.get("reason") == "promo_image_removed")

    run_report = {
        "generated_at": datetime.now().isoformat(),
        "mode": "final_photo_fix_manual_rules",
        "manual_photo_fix_mode": True,
        "source_catalog": str(SRC_DIR / "doors_catalog.csv"),
        "total_rows": len(rows_sorted),
        "producers_count": producers_total,
        **coverage,
        "rows_with_image": rows_with_image,
        "rows_without_image": rows_without_image,
        "material_filled_count": material_filled,
        "needs_review_count": needs_review,
        "rows_by_producer": dict(sorted(Counter(r.get("producer", "") for r in rows_sorted if r.get("producer", "")).items())),
        "photo_fix_summary": {r["producer"]: r for r in photo_summary_rows},
        "entro_model_row_images_used": entro_model_row_used,
        "asd_badge_images_removed": asd_badge_removed,
        "hausdoors_fallback_images_cleared": haus_fallback_cleared,
        "olymp_promo_images_removed": olymp_promo_removed,
    }
    (OUT_DIR / "run_report.json").write_text(json.dumps(run_report, ensure_ascii=False, indent=2), encoding="utf-8")

    print("manual rules build completed")
    print(f"rows={len(rows_sorted)} with_image={rows_with_image} needs_review={needs_review}")
    for r in photo_summary_rows:
        print(
            f"{r['producer']}: replaced={r['replaced_images_count']} removed={r['removed_bad_images_count']} "
            f"empty={r['image_not_found_count']} suspicious={r['suspicious_images_remaining_count']}"
        )


if __name__ == "__main__":
    main()
