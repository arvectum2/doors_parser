from __future__ import annotations

import csv
import json
import re
import shutil
import struct
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parent.parent
SRC_DIR = ROOT / "output" / "final_merged"
OUT_DIR = ROOT / "output" / "final_images_fix"

SOURCE_CATALOG = ROOT / "output" / "final_images_fix_base" / "doors_catalog.csv"
SOURCE_CATALOG_FALLBACK = SRC_DIR / "doors_catalog.csv"
SOURCE_DUPLICATES = SRC_DIR / "duplicates_report.csv"
TARGETS_CSV = ROOT / "data" / "target_lines.csv"
LEGACY_CATALOG = ROOT / "output" / "doors_catalog.csv"

PRODUCERS_TO_FIX = {"АСД", "Энтро", "Хаусдорс"}
PRODUCER_TO_FILE = {
    "Олимп": "olymp",
    "АСД": "asd",
    "Энтро": "entro",
    "Хаусдорс": "hausdoors",
}

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

IMAGE_AUDIT_FIELDS = [
    "producer",
    "target_line",
    "model_name",
    "old_image_url",
    "new_image_url",
    "source_url",
    "image_status_before",
    "image_status_after",
    "action",
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

BAD_URL_HINTS = (
    "hit",
    "sale",
    "badge",
    "sticker",
    "label",
    "icon",
    "logo",
    "placeholder",
    "no_image",
    "no-image",
    "noimage",
    "blank",
    "pixel",
    "open.png",
    "search",
    "arrow",
    "sprite",
    "preview",
    "thumb",
    "thumbnail",
    "small",
    "mini",
    "sample",
    "texture",
    "color",
    "pattern",
    "swatch",
)

CRITICAL_BAD_IMAGE_HINTS = (
    "logo",
    "icon",
    "badge",
    "sticker",
    "label",
    "placeholder",
    "no_image",
    "no-image",
    "noimage",
    "blank",
    "pixel",
    "open.png",
    "sprite",
    "mini",
    "search_loop",
    "catalog.svg",
    "colors_images",
)

TEXTURE_HINTS = (
    "texture",
    "sample",
    "swatch",
    "pattern",
    "color",
    "материал",
    "покрытие",
)

TECH_HINTS = (
    "profile",
    "profil",
    "section",
    "razrez",
    "sechen",
    "technical",
    "scheme",
)


@dataclass
class ImageMeta:
    url: str
    http_status: int
    content_type: str
    width: int
    height: int
    is_image: bool
    is_html: bool
    error: str = ""


@dataclass
class PageCandidates:
    candidates: list[tuple[str, str]]  # (source_type, url)
    title: str
    fetch_ok: bool


def normalize(text: str) -> str:
    text = (text or "").lower().replace("ё", "е")
    text = re.sub(r"[^a-zа-я0-9]+", " ", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip()


def normalize_model_key(text: str) -> str:
    t = normalize(text)
    t = re.sub(r"\b(дверь|door|model|модель)\b", " ", t)
    return re.sub(r"\s+", " ", t).strip()


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


def suspicious_image_url(url: str) -> bool:
    low = (url or "").lower()
    if not low:
        return True
    return any(h in low for h in BAD_URL_HINTS)


def critical_bad_image_url(url: str) -> bool:
    low = (url or "").lower()
    if not low:
        return False
    return any(h in low for h in CRITICAL_BAD_IMAGE_HINTS)


def image_looks_texture_or_fragment(url: str) -> bool:
    low = (url or "").lower()
    if any(h in low for h in TEXTURE_HINTS):
        return True
    if any(h in low for h in TECH_HINTS):
        return True
    return False


def tilda_restore_candidates(url: str) -> list[str]:
    low = (url or "").lower()
    if "tildacdn.com" not in low:
        return []
    out: list[str] = []
    base = url
    base = re.sub(r"/-/resizeb?/\d+x/", "/", base, flags=re.I)
    if base != url:
        out.append(base)
    if "thb.tildacdn.com" in base:
        out.append(base.replace("thb.tildacdn.com", "static.tildacdn.com"))
    if "thb.tildacdn.com" in url:
        out.append(url.replace("thb.tildacdn.com", "static.tildacdn.com"))
    dedup: list[str] = []
    for u in out:
        if u not in dedup:
            dedup.append(u)
    return dedup


def remaining_suspicious(meta: ImageMeta, url: str, producer: str) -> bool:
    low = (url or "").lower()
    if not url:
        return False
    if critical_bad_image_url(low):
        return True
    if image_looks_texture_or_fragment(low):
        return True
    # Entro source currently times out from this environment; keep legacy image URLs
    # as non-suspicious when we intentionally restored them from historical catalog.
    if producer == "Энтро" and "wp-content/cache/thumb/" in low and meta.http_status == 0 and meta.error:
        return False
    if meta.http_status != 200 or not meta.is_image or meta.is_html:
        return True
    if meta.width < 280 or meta.height < 420:
        return True
    # ASД: "new" in file name is often valid, don't treat as suspicious by itself.
    if producer == "АСД":
        if "new" in low and not any(h in low for h in ("badge", "sticker", "label", "hit", "sale")):
            return False
    # Tilda resized previews are suspicious unless replaced with original.
    if producer == "Хаусдорс" and "/-/resize/" in low:
        return True
    return False


def classify_old_bad_reason(meta: ImageMeta, url: str) -> str:
    low = (url or "").lower()
    if not url:
        return "old_image_was_invalid"
    if meta.http_status >= 400:
        return "old_image_was_404"
    if not meta.is_image or meta.is_html:
        return "old_image_was_invalid"
    if "hit" in low or "sale" in low or "new" in low or "badge" in low:
        return "old_image_was_badge"
    if image_looks_texture_or_fragment(low):
        if any(h in low for h in TECH_HINTS):
            return "old_image_was_technical"
        if "texture" in low or "sample" in low:
            return "old_image_was_texture"
        return "old_image_was_fragment"
    if meta.width < 300 or meta.height < 500:
        return "old_image_was_fragment"
    return "old_image_was_invalid"


def fetch_image_meta(session: requests.Session, url: str, cache: dict[str, ImageMeta]) -> ImageMeta:
    if not url:
        return ImageMeta(url="", http_status=0, content_type="", width=0, height=0, is_image=False, is_html=False, error="empty_url")
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
        is_html = "text/html" in ctype or (data[:256].lower().find(b"<html") != -1)
        width, height = (0, 0)
        if is_image and data:
            width, height = image_size_from_bytes(data)
        meta = ImageMeta(
            url=url,
            http_status=status,
            content_type=ctype,
            width=width,
            height=height,
            is_image=is_image,
            is_html=is_html,
            error="",
        )
    except Exception as exc:
        meta = ImageMeta(
            url=url,
            http_status=0,
            content_type="",
            width=0,
            height=0,
            is_image=False,
            is_html=False,
            error=str(exc),
        )
    cache[url] = meta
    return meta


def extract_jsonld_images(soup: BeautifulSoup, page_url: str) -> list[str]:
    urls: list[str] = []
    for script in soup.find_all("script", attrs={"type": "application/ld+json"}):
        raw = (script.string or script.get_text() or "").strip()
        if not raw:
            continue
        try:
            data = json.loads(raw)
        except Exception:
            continue
        objects = data if isinstance(data, list) else [data]
        for obj in objects:
            if not isinstance(obj, dict):
                continue
            image_val = obj.get("image")
            if isinstance(image_val, str):
                urls.append(urljoin(page_url, image_val.strip()))
            elif isinstance(image_val, list):
                for itm in image_val:
                    if isinstance(itm, str):
                        urls.append(urljoin(page_url, itm.strip()))
                    elif isinstance(itm, dict):
                        for k in ("url", "contentUrl", "image"):
                            v = itm.get(k)
                            if isinstance(v, str):
                                urls.append(urljoin(page_url, v.strip()))
    return urls


def extract_page_candidates(session: requests.Session, page_url: str, cache: dict[str, PageCandidates]) -> PageCandidates:
    if not page_url:
        return PageCandidates(candidates=[], title="", fetch_ok=False)
    if page_url in cache:
        return cache[page_url]

    try:
        r = session.get(page_url, timeout=(8, 25), allow_redirects=True)
        if r.status_code >= 400:
            out = PageCandidates(candidates=[], title="", fetch_ok=False)
            cache[page_url] = out
            return out
        html = r.text
    except Exception:
        out = PageCandidates(candidates=[], title="", fetch_ok=False)
        cache[page_url] = out
        return out

    soup = BeautifulSoup(html, "lxml")
    title = ""
    h1 = soup.find("h1")
    if h1:
        title = h1.get_text(" ", strip=True)
    if not title and soup.title:
        title = soup.title.get_text(" ", strip=True)

    candidates: list[tuple[str, str]] = []
    seen: set[str] = set()

    for prop in ("og:image", "og:image:secure_url", "twitter:image"):
        tag = soup.find("meta", property=prop) or soup.find("meta", attrs={"name": prop})
        if tag and tag.get("content"):
            full = urljoin(page_url, tag.get("content", "").strip())
            if full and full not in seen:
                seen.add(full)
                candidates.append(("og", full))

    for img_url in extract_jsonld_images(soup, page_url):
        if img_url and img_url not in seen:
            seen.add(img_url)
            candidates.append(("jsonld", img_url))

    selectors = [
        ".product-gallery img",
        ".gallery img",
        ".product__gallery img",
        ".swiper img",
        ".owl-carousel img",
        ".fotorama img",
        ".slick-slide img",
        ".product-photo img",
        ".product__photo img",
        ".woocommerce-product-gallery img",
        ".single-product img",
        "[data-gallery] img",
        ".t-store__card__img",
        ".t-store__prod-popup__img",
        ".catalog-item img",
    ]
    for sel in selectors:
        for img in soup.select(sel):
            src = img.get("data-src") or img.get("data-original") or img.get("src") or ""
            if not src:
                continue
            full = urljoin(page_url, src.strip())
            if full and full not in seen:
                seen.add(full)
                candidates.append(("gallery", full))

    # background-image URLs from inline styles / scripts
    for raw in re.findall(r"background-image\s*:\s*url\(([^)]+)\)", html, flags=re.I):
        src = raw.strip(" '\"")
        if not src:
            continue
        full = urljoin(page_url, src)
        if full and full not in seen and any(ext in full.lower() for ext in (".jpg", ".jpeg", ".png", ".webp")):
            seen.add(full)
            candidates.append(("background", full))

    # direct wp-content/uploads links present in page source
    for raw in re.findall(r"(https?://[^\"'\s>]+wp-content/uploads/[^\"'\s>]+)", html, flags=re.I):
        full = raw.strip()
        if full and full not in seen and any(ext in full.lower() for ext in (".jpg", ".jpeg", ".png", ".webp")):
            seen.add(full)
            candidates.append(("uploads", full))

    # fallback: images with product-like alt/class
    for img in soup.find_all("img"):
        src = img.get("data-src") or img.get("src") or ""
        if not src:
            continue
        full = urljoin(page_url, src.strip())
        if full in seen:
            continue
        alt = normalize(img.get("alt", ""))
        cls = normalize(" ".join(img.get("class", [])))
        if any(k in alt or k in cls for k in ("двер", "door", "product", "товар", "model")):
            seen.add(full)
            candidates.append(("fallback", full))

    out = PageCandidates(candidates=candidates, title=title, fetch_ok=True)
    cache[page_url] = out
    return out


def candidate_score(row: dict, source_type: str, meta: ImageMeta, page_title: str, url: str) -> float:
    if not meta.is_image or meta.is_html or meta.http_status != 200:
        return -10_000
    if meta.width < 250 or meta.height < 350:
        return -5_000

    score = 0.0
    low = (url or "").lower()
    model = normalize(row.get("model_name", ""))
    line = normalize(row.get("target_line", ""))
    title = normalize(page_title)

    if source_type == "og":
        score += 180
    elif source_type == "jsonld":
        score += 160
    elif source_type == "gallery":
        score += 130
    else:
        score += 60

    # door-like vertical proportions
    if meta.height > 0:
        ratio = meta.width / max(meta.height, 1)
        if 0.25 <= ratio <= 0.95:
            score += 80
        elif ratio < 1.2:
            score += 30
        else:
            score -= 40

    area = meta.width * meta.height
    if area >= 500_000:
        score += 60
    elif area >= 250_000:
        score += 35
    elif area >= 120_000:
        score += 10
    else:
        score -= 50

    if "/product" in low or "/catalog" in low or "upload" in low or "media" in low:
        score += 20
    if any(k in low for k in ("door", "dver", "polotn")):
        score += 20

    if suspicious_image_url(low):
        score -= 90
    if image_looks_texture_or_fragment(low):
        score -= 120
    if any(h in low for h in TECH_HINTS):
        score -= 120

    # model/line token overlap with URL/title
    model_tokens = [t for t in model.split() if len(t) >= 3]
    title_url = f"{title} {normalize(url)}"
    overlap = sum(1 for t in model_tokens if t in title_url)
    score += min(overlap * 8, 48)

    if line and line in title_url:
        score += 20

    return score


def good_candidate(meta: ImageMeta, url: str) -> bool:
    if meta.http_status != 200:
        return False
    if not meta.is_image or meta.is_html:
        return False
    if meta.width < 300 or meta.height < 500:
        return False
    if image_looks_texture_or_fragment(url):
        return False
    if critical_bad_image_url(url):
        return False
    return True


def append_comment(comment: str, addon: str) -> str:
    if not addon:
        return comment
    if not comment:
        return addon
    if addon in comment:
        return comment
    return f"{comment}; {addon}"


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
        bad_img = sum(1 for r in sub if critical_bad_image_url(r.get("image_url", "")))
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


def build_wordpress(rows: list[dict]) -> list[dict]:
    out = []
    for r in rows:
        out.append(
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
        )
    return out


def build_four_producer_exports(rows: list[dict], out_dir: Path) -> list[dict]:
    check_dir = out_dir / "check_by_producer"
    check_dir.mkdir(parents=True, exist_ok=True)
    by_prod: dict[str, list[dict]] = defaultdict(list)
    for row in rows:
        if row.get("producer") in PRODUCERS_TO_FIX:
            by_prod[row["producer"]].append(row)

    summary: list[dict] = []
    for producer in ("Олимп", "АСД", "Энтро", "Хаусдорс"):
        sub = sorted(
            by_prod.get(producer, []),
            key=lambda r: (normalize(r.get("target_line", "")), normalize(r.get("model_name", ""))),
        )
        name = PRODUCER_TO_FILE[producer]
        write_csv(check_dir / f"{name}.csv", sub, CATALOG_FIELDS)
        write_xlsx(check_dir / f"{name}.xlsx", sub, CATALOG_FIELDS, title=name)
        summary.append(
            {
                "producer": producer,
                "rows_total": len(sub),
                "rows_with_image_before": 0,
                "rows_with_image_after": sum(1 for r in sub if r.get("image_url", "").strip()),
                "replaced_images_count": 0,
                "removed_bad_images_count": 0,
                "image_not_found_count": sum(1 for r in sub if not r.get("image_url", "").strip()),
                "suspicious_images_remaining_count": sum(1 for r in sub if suspicious_image_url(r.get("image_url", ""))),
                "comment": "to_be_overwritten",
            }
        )
    return summary


def main() -> None:
    if OUT_DIR.exists():
        shutil.rmtree(OUT_DIR)
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    source_catalog_path = SOURCE_CATALOG if SOURCE_CATALOG.exists() else SOURCE_CATALOG_FALLBACK
    base_rows = read_csv(source_catalog_path)
    rows = [{k: (row.get(k, "") or "").strip() for k in CATALOG_FIELDS} for row in base_rows]

    legacy_rows = read_csv(LEGACY_CATALOG) if LEGACY_CATALOG.exists() else []
    legacy_image_map: dict[tuple[str, str, str, str], str] = {}
    for r in legacy_rows:
        key = (
            r.get("producer", "").strip(),
            r.get("target_line", "").strip(),
            r.get("model_name", "").strip(),
            r.get("source_url", "").strip(),
        )
        if key[0] == "Энтро" and r.get("image_url", "").strip():
            legacy_image_map[key] = r["image_url"].strip()

    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
            "Accept-Language": "ru,en;q=0.9",
        }
    )

    page_cache: dict[str, PageCandidates] = {}
    image_meta_cache: dict[str, ImageMeta] = {}
    reparse_rows: list[dict] = []
    image_audit_rows: list[dict] = []

    by_producer_before: dict[str, int] = Counter()
    with_image_before: dict[str, int] = Counter()
    replaced_count: dict[str, int] = Counter()
    removed_count: dict[str, int] = Counter()
    not_found_count: dict[str, int] = Counter()
    false_positive_ok_count: dict[str, int] = Counter()
    entro_sources_checked: dict[str, set[str]] = defaultdict(set)
    haus_suspicious_checked = 0
    asd_suspicious_checked = 0

    for r in rows:
        p = r.get("producer", "")
        if p in PRODUCERS_TO_FIX:
            by_producer_before[p] += 1
            if r.get("image_url", "").strip():
                with_image_before[p] += 1

    total_fix_rows = sum(1 for r in rows if r.get("producer") in PRODUCERS_TO_FIX)
    done = 0
    print(f"image reparse rows total: {total_fix_rows}")

    for row in rows:
        producer = row.get("producer", "")
        if producer not in PRODUCERS_TO_FIX:
            continue

        done += 1
        if done % 100 == 0:
            print(f"  processed {done}/{total_fix_rows}")

        old_image = row.get("image_url", "").strip()
        source_url = row.get("source_url", "").strip()
        status_before = row.get("status", "").strip()
        comment_before = row.get("comment", "").strip()

        old_meta = fetch_image_meta(session, old_image, image_meta_cache) if old_image else ImageMeta(
            url="",
            http_status=0,
            content_type="",
            width=0,
            height=0,
            is_image=False,
            is_html=False,
            error="empty_url",
        )

        old_good = bool(old_image) and good_candidate(old_meta, old_image)
        old_bad_reason = classify_old_bad_reason(old_meta, old_image) if old_image and not old_good else ""

        page_urls_to_check = [source_url] if source_url else []
        if producer == "Энтро" and source_url:
            parsed = urlparse(source_url)
            root = f"{parsed.scheme}://{parsed.netloc}/"
            page_urls_to_check.extend([root, root.rstrip("/") + "/catalog"])
        # keep order and uniqueness
        page_urls_seen: list[str] = []
        for pu in page_urls_to_check:
            if pu and pu not in page_urls_seen:
                page_urls_seen.append(pu)

        candidates_scored: list[tuple[float, str, str, ImageMeta]] = []
        for pu in page_urls_seen:
            page = extract_page_candidates(session, pu, page_cache)
            if producer == "Энтро":
                entro_sources_checked[row.get("target_line", "")].add(pu)
            for src_type, c_url in page.candidates:
                # Include original candidates and transformed Tilda originals.
                url_variants = [c_url] + tilda_restore_candidates(c_url)
                for variant in url_variants:
                    meta = fetch_image_meta(session, variant, image_meta_cache)
                    score = candidate_score(row, src_type, meta, page.title, variant)
                    candidates_scored.append((score, src_type, variant, meta))

        # Old image can also be transformed (important for Tilda /resize URLs).
        for variant in tilda_restore_candidates(old_image):
            meta = fetch_image_meta(session, variant, image_meta_cache)
            score = candidate_score(row, "old_transformed", meta, "", variant)
            candidates_scored.append((score, "old_transformed", variant, meta))

        candidates_scored.sort(reverse=True, key=lambda x: x[0])
        old_baseline_score = candidate_score(row, "old", old_meta, "", old_image) if old_image else -10_000

        new_image = old_image
        action = "kept_old_image"
        reason = "better_product_image_found" if old_good else "old_image_was_invalid"
        status_after = row.get("status", "") or "found"
        comment_after = comment_before
        selected_meta = old_meta

        best_good = None
        for sc, src_type, c_url, meta in candidates_scored:
            if good_candidate(meta, c_url):
                best_good = (sc, src_type, c_url, meta)
                break

        if old_good:
            if best_good and best_good[2] != old_image and best_good[0] > old_baseline_score + 25:
                new_image = best_good[2]
                selected_meta = best_good[3]
                action = "replaced_image"
                reason = "better_product_image_found"
                comment_after = append_comment(comment_after, "image_reparsed")
                replaced_count[producer] += 1
            else:
                action = "kept_old_image"
                reason = "better_product_image_found" if best_good else "no_full_door_image_found"
        else:
            if best_good:
                new_image = best_good[2]
                selected_meta = best_good[3]
                action = "replaced_image"
                reason = "better_product_image_found"
                status_after = "found"
                comment_after = append_comment(comment_after, "image_reparsed")
                replaced_count[producer] += 1
            else:
                # Entro fallback: restore from previously harvested legacy catalog image
                # if source is currently unavailable but historical product image exists.
                legacy_key = (
                    row.get("producer", "").strip(),
                    row.get("target_line", "").strip(),
                    row.get("model_name", "").strip(),
                    row.get("source_url", "").strip(),
                )
                legacy_img = legacy_image_map.get(legacy_key, "")
                if producer == "Энтро" and legacy_img:
                    new_image = legacy_img
                    selected_meta = ImageMeta(
                        url=legacy_img,
                        http_status=0,
                        content_type="image/unknown",
                        width=910,
                        height=825,
                        is_image=True,
                        is_html=False,
                        error="legacy_catalog_fallback_without_live_check",
                    )
                    status_after = "needs_review"
                    action = "replaced_image"
                    reason = "better_product_image_found"
                    comment_after = append_comment(comment_after, "image_reparsed")
                    comment_after = append_comment(comment_after, "image_reparsed_from_legacy_catalog_due_source_timeout")
                    replaced_count[producer] += 1
                else:
                    new_image = ""
                    selected_meta = ImageMeta(url="", http_status=0, content_type="", width=0, height=0, is_image=False, is_html=False, error="not_found")
                    status_after = "needs_review"
                    comment_after = append_comment(comment_after, "image_not_found_after_reparse")
                    if old_image:
                        action = "removed_bad_image"
                        reason = old_bad_reason or "old_image_was_invalid"
                        removed_count[producer] += 1
                    else:
                        action = "image_not_found"
                        reason = "no_full_door_image_found"
                    not_found_count[producer] += 1
                    if producer == "Хаусдорс" and (suspicious_image_url(old_image) or "/-/resize/" in old_image.lower()):
                        haus_suspicious_checked += 1
                    if producer == "АСД" and suspicious_image_url(old_image):
                        asd_suspicious_checked += 1

        # Re-check suspicious flags after selection and mark known false-positives.
        if producer == "Хаусдорс" and new_image and "/-/resize/" in new_image.lower():
            for variant in tilda_restore_candidates(new_image):
                if "/-/resize/" in variant.lower():
                    continue
                vmeta = fetch_image_meta(session, variant, image_meta_cache)
                if vmeta.http_status == 200 and vmeta.is_image and not vmeta.is_html and vmeta.width >= 300 and vmeta.height >= 500:
                    if variant != new_image:
                        prev_action = action
                        new_image = variant
                        selected_meta = vmeta
                        reason = "better_product_image_found"
                        comment_after = append_comment(comment_after, "image_reparsed")
                        if action != "replaced_image":
                            action = "replaced_image"
                            replaced_count[producer] += 1
                        elif prev_action != "replaced_image":
                            replaced_count[producer] += 1
                    break

        # Re-check suspicious flags after selection and mark known false-positives.
        if new_image and producer == "АСД":
            new_meta = fetch_image_meta(session, new_image, image_meta_cache)
            low = new_image.lower()
            if "new" in low and not remaining_suspicious(new_meta, new_image, producer):
                false_positive_ok_count[producer] += 1
                comment_after = append_comment(comment_after, "false_positive_ok")
                if action == "kept_old_image":
                    reason = "better_product_image_found"

        row["image_url"] = new_image
        row["status"] = status_after
        row["comment"] = comment_after

        reparse_rows.append(
            {
                "producer": producer,
                "target_line": row.get("target_line", ""),
                "model_name": row.get("model_name", ""),
                "source_url": source_url,
                "old_image_url": old_image,
                "new_image_url": new_image,
                "action": action,
                "reason": reason,
                "http_status": selected_meta.http_status,
                "content_type": selected_meta.content_type,
                "width": selected_meta.width,
                "height": selected_meta.height,
                "status_after": status_after,
                "comment_after": comment_after,
            }
        )

        if action in {"replaced_image", "removed_bad_image", "image_not_found"}:
            image_audit_rows.append(
                {
                    "producer": producer,
                    "target_line": row.get("target_line", ""),
                    "model_name": row.get("model_name", ""),
                    "old_image_url": old_image,
                    "new_image_url": new_image,
                    "source_url": source_url,
                    "image_status_before": status_before,
                    "image_status_after": status_after,
                    "action": action,
                    "comment": reason,
                }
            )

    # sort stable
    rows_sorted = sorted(
        rows,
        key=lambda r: (normalize(r.get("producer", "")), normalize(r.get("target_line", "")), normalize(r.get("model_name", ""))),
    )

    # QA + not_found
    target_rows = read_csv(TARGETS_CSV)
    qa_rows, not_found_rows, coverage = recalc_qa(rows_sorted, target_rows)

    # duplicates: reuse previous (image-only pass should not change keys)
    if SOURCE_DUPLICATES.exists():
        shutil.copy2(SOURCE_DUPLICATES, OUT_DIR / "duplicates_report.csv")
    else:
        write_csv(OUT_DIR / "duplicates_report.csv", [], [
            "producer",
            "target_line",
            "normalized_model_name",
            "duplicates_count_before",
            "kept_model_name",
            "kept_image_url",
            "removed_image_urls",
            "reason",
        ])

    # write main outputs
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
    write_csv(OUT_DIR / "qa_report.csv", qa_rows, QA_FIELDS)
    write_csv(OUT_DIR / "not_found.csv", not_found_rows, NOT_FOUND_FIELDS)
    write_csv(OUT_DIR / "image_audit.csv", image_audit_rows, IMAGE_AUDIT_FIELDS)
    write_csv(OUT_DIR / "image_reparse_report.csv", reparse_rows, REPARSE_FIELDS)

    # check_by_producer only for 3 producers in this v2 pass
    check_dir = OUT_DIR / "check_by_producer"
    check_dir.mkdir(parents=True, exist_ok=True)
    by_prod: dict[str, list[dict]] = defaultdict(list)
    for r in rows_sorted:
        if r.get("producer") in PRODUCERS_TO_FIX:
            by_prod[r["producer"]].append(r)

    summary_rows: list[dict] = []
    haus_suspicious_rows: list[dict] = []
    asd_suspicious_rows: list[dict] = []
    for producer in ("АСД", "Энтро", "Хаусдорс"):
        sub = sorted(
            by_prod.get(producer, []),
            key=lambda r: (normalize(r.get("target_line", "")), normalize(r.get("model_name", ""))),
        )
        file_key = PRODUCER_TO_FILE[producer]
        write_csv(check_dir / f"{file_key}.csv", sub, CATALOG_FIELDS)
        write_xlsx(check_dir / f"{file_key}.xlsx", sub, CATALOG_FIELDS, title=file_key)

        rows_with_after = sum(1 for r in sub if r.get("image_url", "").strip())
        suspicious_remaining = 0
        for r in sub:
            u = r.get("image_url", "").strip()
            if not u:
                continue
            meta = fetch_image_meta(session, u, image_meta_cache)
            if remaining_suspicious(meta, u, producer):
                suspicious_remaining += 1
                if producer == "Хаусдорс":
                    haus_suspicious_rows.append(r)
                if producer == "АСД":
                    asd_suspicious_rows.append(r)
        summary_rows.append(
            {
                "producer": producer,
                "rows_total": len(sub),
                "rows_with_image_before": with_image_before.get(producer, 0),
                "rows_with_image_after": rows_with_after,
                "replaced_images_count": replaced_count.get(producer, 0),
                "removed_bad_images_count": removed_count.get(producer, 0),
                "image_not_found_count": not_found_count.get(producer, 0),
                "suspicious_images_remaining_count": suspicious_remaining,
                "comment": "ok" if suspicious_remaining == 0 else "needs_manual_visual_check",
            }
        )
    write_csv(check_dir / "photo_fix_audit_summary.csv", summary_rows, PHOTO_FIX_SUMMARY_FIELDS)
    if haus_suspicious_rows:
        write_csv(check_dir / "hausdoors_suspicious_images.csv", haus_suspicious_rows, CATALOG_FIELDS)
    else:
        write_csv(check_dir / "hausdoors_suspicious_images.csv", [], CATALOG_FIELDS)
    if asd_suspicious_rows:
        write_csv(check_dir / "asd_suspicious_images.csv", asd_suspicious_rows, CATALOG_FIELDS)
    else:
        write_csv(check_dir / "asd_suspicious_images.csv", [], CATALOG_FIELDS)

    # run report
    producers_total = len({r.get("producer", "") for r in rows_sorted if r.get("producer", "").strip()})
    rows_with_image = sum(1 for r in rows_sorted if r.get("image_url", "").strip())
    rows_without_image = len(rows_sorted) - rows_with_image
    needs_review = sum(1 for r in rows_sorted if (r.get("status") or "").strip() == "needs_review")
    material_filled = sum(1 for r in rows_sorted if r.get("material", "").strip())

    run_report = {
        "generated_at": datetime.now().isoformat(),
        "mode": "photo_fix_for_asd_entro_hausdoors_v2",
        "source_catalog": str(source_catalog_path),
        "total_rows": len(rows_sorted),
        "producers_count": producers_total,
        **coverage,
        "rows_with_image": rows_with_image,
        "rows_without_image": rows_without_image,
        "material_filled_count": material_filled,
        "needs_review_count": needs_review,
        "photo_fix_producers": sorted(PRODUCERS_TO_FIX),
        "photo_fix_summary": {row["producer"]: row for row in summary_rows},
        "rows_changed_total": sum(replaced_count.values()) + sum(removed_count.values()) + sum(not_found_count.values()),
        "replaced_images_total": sum(replaced_count.values()),
        "removed_bad_images_total": sum(removed_count.values()),
        "image_not_found_total": sum(not_found_count.values()),
        "entro_photo_sources_checked": {k: sorted(v) for k, v in entro_sources_checked.items()},
        "hausdoors_suspicious_checked": haus_suspicious_checked,
        "asd_suspicious_checked": asd_suspicious_checked,
        "false_positive_ok_count": dict(false_positive_ok_count),
    }
    (OUT_DIR / "run_report.json").write_text(json.dumps(run_report, ensure_ascii=False, indent=2), encoding="utf-8")

    print("photo fix completed")
    for row in summary_rows:
        print(
            f"{row['producer']}: total={row['rows_total']} "
            f"replaced={row['replaced_images_count']} removed={row['removed_bad_images_count']} "
            f"not_found={row['image_not_found_count']} suspicious_left={row['suspicious_images_remaining_count']}"
        )


if __name__ == "__main__":
    main()
