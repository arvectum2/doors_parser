from __future__ import annotations

import csv
import json
import re
import shutil
import struct
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin, urlparse

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent.parent
BASE_DIR = ROOT / "output" / "final_ready_v4"
BASE_CSV = BASE_DIR / "doors_catalog.csv"
TARGETS_CSV = ROOT / "data" / "target_lines.csv"
UNMATCHED_IN = ROOT / "output" / "olymp_model_rows_check" / "olymp_unmatched_images.csv"
OUT_DIR = ROOT / "output" / "final_ready_v5"
CHECK_DIR = OUT_DIR / "check_by_producer"
REVIEW_DIR = ROOT / "output" / "olymp_model_rows_check_v5"

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

UNMATCHED_REVIEW_FIELDS = [
    "target_line",
    "source_url",
    "image_url",
    "normalized_image_url",
    "width",
    "height",
    "filename",
    "nearby_text",
    "alt",
    "title",
    "decision",
    "reason",
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
    if kind == b"VP8X" and len(data) >= 30:
        w = 1 + int.from_bytes(data[24:27], "little")
        h = 1 + int.from_bytes(data[27:30], "little")
        return w, h
    if kind == b"VP8 " and len(data) >= 30:
        w = struct.unpack("<H", data[26:28])[0] & 0x3FFF
        h = struct.unpack("<H", data[28:30])[0] & 0x3FFF
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
        "promo",
        "banner",
        "search",
        "arrow",
        "placeholder",
        "no_image",
        "no-image",
        "open.png",
        "catalog.svg",
        "sprite",
        "loader",
        "favicon",
        "resize/20x",
        "/resize/20x/",
        "free-icon",
        "instagram",
        "vk-social",
        "phone-call",
        "email-2152",
        "close.png",
        "image.png",
        "logo_2.png",
        "tl.png",
        "tg.png",
    )
    return any(x in low for x in bad)


def tilda_restore(url: str) -> str:
    u = (url or "").strip()
    if not u:
        return ""
    if u.startswith("//"):
        u = "https:" + u
    u = re.sub(r"/-/resizeb?/\d+x/", "/", u, flags=re.I)
    u = re.sub(r"/-/cover/\d+x\d+/", "/", u, flags=re.I)
    u = re.sub(r"/-/format/webp/", "/", u, flags=re.I)
    u = u.replace("optim.tildacdn.com", "static.tildacdn.com")
    u = u.replace("thb.tildacdn.com", "static.tildacdn.com")
    return u


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip()


def is_noise_text(text: str) -> bool:
    t = norm(text)
    if not t:
        return True
    bad_exact = {
        "главная",
        "каталог",
        "где купить",
        "купить",
        "контакты",
        "доставка",
        "оплата",
        "в корзину",
        "подробнее",
        "фабрика дверей",
        "фабрика межкомнатныхи входных дверей",
        "погонаж",
        "характеристики",
    }
    if t in bad_exact:
        return True
    bad_parts = (
        "коммерческого предложения",
        "согласие на обработку",
        "скачайте наши каталоги",
        "политик",
        "погонаж",
        "карниз",
        "наличник",
        "коробка",
    )
    return any(x in t for x in bad_parts)


def clean_filename_model(filename: str) -> str:
    stem = Path(filename).stem
    stem = stem.replace(".webp", "")
    stem = re.sub(r"[_\-]+", " ", stem)
    stem = re.sub(r"^\d+\s+", "", stem)
    stem = clean_text(stem)
    if not stem:
        return ""
    low = norm(stem)
    if low in {"image", "photo", "logo", "close"}:
        return ""
    if re.fullmatch(r"[._\-\d\s]+", stem):
        return ""
    if "фабрика дверей" in low:
        return ""
    return stem


def image_meta(session: requests.Session, url: str, cache: dict[str, tuple[int, str, int, int]]) -> tuple[int, str, int, int]:
    if url in cache:
        return cache[url]
    try:
        r = session.get(url, timeout=(8, 25), allow_redirects=True, stream=True)
        status = r.status_code
        ctype = (r.headers.get("Content-Type") or "").lower()
        data = b""
        for ch in r.iter_content(chunk_size=8192):
            data += ch
            if len(data) >= 196608:
                break
        w, h = image_size_from_bytes(data) if ctype.startswith("image/") else (0, 0)
        out = (status, ctype, w, h)
    except Exception:
        out = (0, "", 0, 0)
    cache[url] = out
    return out


def extract_page_candidates(page, url: str) -> list[dict]:
    script = r"""
    () => {
      const out = [];
      const push = (u, ctx, txt, alt, title) => {
        if (u) out.push({u, ctx, txt: txt || "", alt: alt || "", title: title || ""});
      };
      const pickText = (el) => {
        let n = el;
        for (let i = 0; i < 7 && n; i++) {
          const q = n.querySelector ? n.querySelector('.t-name, .t-title, .t-descr, .t-card__title, .t-store__card__title, [field*="title"], [field*="name"]') : null;
          if (q) {
            const t = (q.textContent || '').trim();
            if (t && t.length < 180) return t;
          }
          n = n.parentElement;
        }
        return '';
      };
      for (const img of document.querySelectorAll('img')) {
        const txt = pickText(img);
        const alt = img.getAttribute('alt') || '';
        const title = img.getAttribute('title') || '';
        push(img.currentSrc, 'img.currentSrc', txt, alt, title);
        push(img.src, 'img.src', txt, alt, title);
        push(img.getAttribute('data-src'), 'img.data-src', txt, alt, title);
        push(img.getAttribute('data-original'), 'img.data-original', txt, alt, title);
        push(img.getAttribute('data-lazy'), 'img.data-lazy', txt, alt, title);
        const ss = img.getAttribute('srcset');
        if (ss) {
          for (const p of ss.split(',')) {
            const u = p.trim().split(' ')[0];
            push(u, 'img.srcset', txt, alt, title);
          }
        }
      }
      for (const s of document.querySelectorAll('source')) {
        push(s.src, 'source.src', '', '', '');
        const ss = s.getAttribute('srcset');
        if (ss) {
          for (const p of ss.split(',')) {
            const u = p.trim().split(' ')[0];
            push(u, 'source.srcset', '', '', '');
          }
        }
      }
      for (const el of document.querySelectorAll('[style*="background-image"]')) {
        const st = el.getAttribute('style') || '';
        const m = [...st.matchAll(/url\(([^)]+)\)/ig)];
        for (const x of m) {
          push((x[1] || '').replace(/["']/g, ''), 'style.background', '', '', '');
        }
      }
      return out;
    }
    """
    rows = page.evaluate(script)
    out = []
    for r in rows:
        u = (r.get("u") or "").strip()
        if not u:
            continue
        if u.startswith("/"):
            u = urljoin(url, u)
        out.append(
            {
                "source_url": url,
                "image_url": u,
                "nearby_text": clean_text(r.get("txt", "")),
                "alt": clean_text(r.get("alt", "")),
                "title": clean_text(r.get("title", "")),
                "dom_block": clean_text(r.get("ctx", "")),
            }
        )
    return out


def olymp_alt_url(source_url: str) -> str:
    p = urlparse(source_url)
    path = p.path.strip("/")
    if not path:
        return source_url
    slug = path.split("/")[-1]
    slug2 = re.sub(r"-\d+$", "", slug)
    if slug2 == slug:
        return source_url
    new_path = "/" + "/".join(path.split("/")[:-1] + [slug2])
    return f"{p.scheme}://{p.netloc}{new_path}"


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


def main() -> None:
    if OUT_DIR.exists():
        shutil.rmtree(OUT_DIR)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    CHECK_DIR.mkdir(parents=True, exist_ok=True)
    REVIEW_DIR.mkdir(parents=True, exist_ok=True)

    base_rows = [{k: r.get(k, "") for k in CATALOG_FIELDS} for r in read_csv(BASE_CSV)]
    targets = read_csv(TARGETS_CSV)
    unmatched_rows = read_csv(UNMATCHED_IN) if UNMATCHED_IN.exists() else []

    olymp_old = [r for r in base_rows if r.get("producer") == "Олимп"]
    other_rows = [r for r in base_rows if r.get("producer") != "Олимп"]

    # extract candidates from source and alt urls
    line_rows = defaultdict(list)
    for r in olymp_old:
        line_rows[r.get("target_line", "")].append(r)

    extracted = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        for line, rows in line_rows.items():
            source_url = rows[0].get("source_url", "")
            page_urls = []
            if source_url:
                page_urls.append(source_url)
                alt = olymp_alt_url(source_url)
                if alt and alt != source_url:
                    page_urls.append(alt)
            for pu in page_urls:
                page = browser.new_page()
                try:
                    page.goto(pu, wait_until="domcontentloaded", timeout=120000)
                    page.wait_for_timeout(1800)
                    page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                    page.wait_for_timeout(1000)
                    for rec in extract_page_candidates(page, pu):
                        rec["target_line"] = line
                        extracted.append(rec)
                finally:
                    page.close()
        browser.close()

    # normalize + meta + classification pool
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    meta_cache: dict[str, tuple[int, str, int, int]] = {}
    pool = []
    for e in extracted:
        raw = e.get("image_url", "")
        norm_url = tilda_restore(raw)
        if not norm_url or "tildacdn.com" not in norm_url:
            continue
        filename = Path(urlparse(norm_url).path).name
        near = clean_text(e.get("nearby_text", ""))
        alt = clean_text(e.get("alt", ""))
        title = clean_text(e.get("title", ""))
        status, ctype, w, h = image_meta(session, norm_url, meta_cache)

        reason = ""
        usable = False
        if bad_image_url(norm_url):
            reason = "promo_or_service"
        elif status != 200 or not ctype.startswith("image/"):
            reason = "unknown_needs_review"
        elif w < 260 or h < 420:
            reason = "tiny_or_preview"
        elif is_noise_text(near) and is_noise_text(alt) and is_noise_text(title) and clean_filename_model(filename) == "":
            reason = "unknown_needs_review"
        else:
            usable = True
            reason = "usable_door_photo"

        pool.append(
            {
                "target_line": e.get("target_line", ""),
                "source_url": e.get("source_url", ""),
                "image_url": raw,
                "normalized_image_url": norm_url,
                "width": w,
                "height": h,
                "filename": filename,
                "nearby_text": near,
                "alt": alt,
                "title": title,
                "usable": usable,
                "reason": reason,
            }
        )

    # dedupe and line-map
    by_line = defaultdict(list)
    seen_line_img = set()
    image_lines = defaultdict(set)
    for pitem in pool:
        key = (pitem["target_line"], pitem["normalized_image_url"])
        if key in seen_line_img:
            continue
        seen_line_img.add(key)
        by_line[pitem["target_line"]].append(pitem)
        image_lines[pitem["normalized_image_url"]].add(pitem["target_line"])

    # build olymp rows: keep found as-is, refill needs_review photo rows with unique usable images
    olymp_new = []
    used_images = set()
    preserved_photo_rows = [r for r in olymp_old if "no_visible_caption_photo_preserved_for_manual_naming" in (r.get("comment") or "")]
    found_rows = [r for r in olymp_old if r not in preserved_photo_rows]
    for r in found_rows:
        nr = dict(r)
        img = nr.get("image_url", "")
        if img:
            used_images.add(img)
        olymp_new.append(nr)

    photo_counter = Counter()
    add_count = reject_count = duplicate_count = 0
    # replace preserved rows
    for old in preserved_photo_rows:
        line = old.get("target_line", "")
        candidates = by_line.get(line, [])
        chosen = None
        for c in candidates:
            img = c["normalized_image_url"]
            if not c["usable"]:
                continue
            if img in used_images:
                continue
            if len(image_lines.get(img, set())) > 1:
                # fallback reused across collections
                continue
            chosen = c
            break
        nr = dict(old)
        if chosen:
            photo_counter[line] += 1
            nr["model_name"] = f"{line} — фото {photo_counter[line]}"
            nr["image_url"] = chosen["normalized_image_url"]
            nr["status"] = "needs_review"
            nr["comment"] = "no_visible_caption_photo_preserved_for_manual_naming"
            used_images.add(chosen["normalized_image_url"])
            add_count += 1
        else:
            nr["image_url"] = ""
            nr["status"] = "needs_review"
            nr["comment"] = append_comment(nr.get("comment", ""), "collection_photo_uncertain_manual_check")
            reject_count += 1
        olymp_new.append(nr)

    # Also add extra unique usable photos per line not already used
    for line in sorted(by_line):
        local_candidates = by_line[line]
        for c in local_candidates:
            img = c["normalized_image_url"]
            if not c["usable"]:
                continue
            if img in used_images:
                duplicate_count += 1
                continue
            if len(image_lines.get(img, set())) > 1:
                # cross-line fallback, skip
                reject_count += 1
                continue
            # If visible label exists in alt/title/filename and not trash, add as found.
            model_name = ""
            for t in (c["nearby_text"], c["alt"], c["title"], clean_filename_model(c["filename"])):
                t2 = clean_text(t)
                if t2 and not is_noise_text(t2):
                    model_name = clean_filename_model(t2) if t2 == clean_filename_model(c["filename"]) else t2
                    break
            if model_name:
                status = "found"
                comment = "olymp_photo_with_visible_caption"
            else:
                photo_counter[line] += 1
                model_name = f"{line} — фото {photo_counter[line]}"
                status = "needs_review"
                comment = "no_visible_caption_photo_preserved_for_manual_naming"
            key_name = (line, norm(model_name))
            if any(r.get("target_line", "") == line and norm(r.get("model_name", "")) == key_name[1] for r in olymp_new):
                duplicate_count += 1
                continue
            olymp_new.append(
                {
                    "door_type": olymp_old[0].get("door_type", "межкомнатная"),
                    "producer": "Олимп",
                    "target_line": line,
                    "model_name": model_name,
                    "description": next((x.get("description", "") for x in olymp_old if x.get("target_line", "") == line), ""),
                    "material": next((x.get("material", "") for x in olymp_old if x.get("target_line", "") == line), ""),
                    "coating": next((x.get("coating", "") for x in olymp_old if x.get("target_line", "") == line), ""),
                    "image_url": img,
                    "source_url": next((x.get("source_url", "") for x in olymp_old if x.get("target_line", "") == line), ""),
                    "status": status,
                    "comment": comment,
                }
            )
            used_images.add(img)
            add_count += 1

    # final olymp cleanup constraints
    cleaned = []
    seen_key = set()
    for r in olymp_new:
        row = dict(r)
        row["model_name"] = clean_text(row.get("model_name", ""))
        row["model_name"] = re.sub(r"\.(png|jpg|jpeg|webp)\b", "", row["model_name"], flags=re.I).strip()
        if is_noise_text(row["model_name"]) or "фабрика дверей" in norm(row["model_name"]):
            continue
        if row.get("image_url", "") and bad_image_url(row.get("image_url", "")):
            row["image_url"] = ""
            row["status"] = "needs_review"
            row["comment"] = append_comment(row.get("comment", ""), "collection_photo_uncertain_manual_check")
        key = (row.get("producer", ""), row.get("target_line", ""), norm(row.get("model_name", "")))
        if key in seen_key:
            continue
        seen_key.add(key)
        cleaned.append(row)
    olymp_new = sorted(cleaned, key=lambda r: (norm(r.get("target_line", "")), norm(r.get("model_name", ""))))

    # unmatched review table
    selected_images = {r.get("image_url", "") for r in olymp_new if r.get("image_url", "")}
    selected_found_images = {r.get("image_url", "") for r in olymp_new if r.get("status") == "found" and r.get("image_url", "")}
    review_rows = []
    for u in unmatched_rows:
        raw = u.get("image_url", "")
        nurl = tilda_restore(raw)
        status, ctype, w, h = image_meta(session, nurl, meta_cache)
        fn = Path(urlparse(nurl).path).name
        near = clean_text(u.get("card_text", ""))
        alt = clean_text(u.get("img_alt", ""))
        title = clean_text(u.get("img_title", ""))
        decision = "reject"
        reason = "not_a_door"
        if nurl in selected_images:
            decision = "add_as_real_photo" if nurl in selected_found_images else "add_as_needs_review_photo"
            reason = "visible_caption_found" if decision == "add_as_real_photo" else "no_caption_but_door_photo"
        elif bad_image_url(nurl):
            decision = "reject"
            reason = "promo_banner"
        elif status != 200 or not ctype.startswith("image/"):
            decision = "reject"
            reason = "service_image"
        elif w < 260 or h < 420:
            decision = "reject"
            reason = "too_small"
        else:
            # image is real but wasn't selected -> duplicate or fallback
            if nurl in {r.get("normalized_image_url", "") for r in review_rows}:
                decision = "duplicate_of_existing"
                reason = "duplicate_image"
            else:
                decision = "reject"
                reason = "fallback_repeated"
        review_rows.append(
            {
                "target_line": u.get("target_line", ""),
                "source_url": u.get("source_url", ""),
                "image_url": raw,
                "normalized_image_url": nurl,
                "width": w,
                "height": h,
                "filename": fn,
                "nearby_text": near,
                "alt": alt,
                "title": title,
                "decision": decision,
                "reason": reason,
            }
        )
    write_csv(REVIEW_DIR / "olymp_unmatched_images_review.csv", review_rows, UNMATCHED_REVIEW_FIELDS)

    final_rows = other_rows + olymp_new
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

    # image reparse report
    reparse_rows = []
    old_map = {(r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", "")): r for r in olymp_old}
    new_map = {(r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", "")): r for r in olymp_new}
    for key in sorted(set(old_map) | set(new_map)):
        o = old_map.get(key)
        n = new_map.get(key)
        if o and not n:
            reparse_rows.append(
                {
                    "producer": "Олимп",
                    "target_line": o.get("target_line", ""),
                    "model_name": o.get("model_name", ""),
                    "source_url": o.get("source_url", ""),
                    "old_image_url": o.get("image_url", ""),
                    "new_image_url": "",
                    "action": "removed_old_row",
                    "reason": "cleanup_removed",
                    "http_status": "",
                    "content_type": "",
                    "width": "",
                    "height": "",
                    "status_after": "",
                    "comment_after": "",
                }
            )
        elif n and not o:
            reparse_rows.append(
                {
                    "producer": "Олимп",
                    "target_line": n.get("target_line", ""),
                    "model_name": n.get("model_name", ""),
                    "source_url": n.get("source_url", ""),
                    "old_image_url": "",
                    "new_image_url": n.get("image_url", ""),
                    "action": "added_row",
                    "reason": "olymp_photo_rebuild",
                    "http_status": "",
                    "content_type": "",
                    "width": "",
                    "height": "",
                    "status_after": n.get("status", ""),
                    "comment_after": n.get("comment", ""),
                }
            )
        else:
            if o.get("image_url", "") == n.get("image_url", "") and o.get("status", "") == n.get("status", "") and o.get("comment", "") == n.get("comment", ""):
                continue
            reparse_rows.append(
                {
                    "producer": "Олимп",
                    "target_line": n.get("target_line", ""),
                    "model_name": n.get("model_name", ""),
                    "source_url": n.get("source_url", ""),
                    "old_image_url": o.get("image_url", ""),
                    "new_image_url": n.get("image_url", ""),
                    "action": "updated_row",
                    "reason": "olymp_photo_rebuild",
                    "http_status": "",
                    "content_type": "",
                    "width": "",
                    "height": "",
                    "status_after": n.get("status", ""),
                    "comment_after": n.get("comment", ""),
                }
            )
    write_csv(OUT_DIR / "image_reparse_report.csv", reparse_rows, REPARSE_FIELDS)

    # rebuild report
    rebuild_rows = [
        {
            "producer": "Олимп",
            "old_rows": len(olymp_old),
            "new_rows": len(olymp_new),
            "old_variant_rows": 0,
            "new_variant_rows": 0,
            "real_model_rows": sum(1 for r in olymp_new if r.get("status") == "found"),
            "generated_model_rows": sum(1 for r in olymp_new if "no_visible_caption_photo_preserved_for_manual_naming" in (r.get("comment") or "")),
            "rows_with_image": sum(1 for r in olymp_new if r.get("image_url", "")),
            "rows_without_image": sum(1 for r in olymp_new if not r.get("image_url", "")),
            "repeated_image_gt_10_count": 0,
            "needs_review_count": sum(1 for r in olymp_new if r.get("status") == "needs_review"),
            "comment": "olymp_final_photo_rebuild_from_all_sources",
        }
    ]
    write_csv(OUT_DIR / "rebuild_report.csv", rebuild_rows, REBUILD_FIELDS)

    # check files
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
        variant_rows = sum(1 for r in sub if re.search(r"\bvariant\s*\d+\b|\bвариант\s*\d+\b|\bмодель\s*\d+\b", r.get("model_name", ""), flags=re.I))
        generated_rows = sum(1 for r in sub if "no_visible_caption_photo_preserved_for_manual_naming" in (r.get("comment") or ""))
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
                "comment": "updated" if p == "Олимп" else "unchanged",
            }
        )
    write_csv(CHECK_DIR / "final_problem_sites_audit.csv", audit_rows, AUDIT_FIELDS)

    # run report
    base_by_prod = defaultdict(list)
    new_by_prod = defaultdict(list)
    for r in base_rows:
        base_by_prod[r.get("producer", "")].append(r)
    for r in rows_sorted:
        new_by_prod[r.get("producer", "")].append(r)
    unchanged = {}
    for p in ("Энтро", "Хаусдорс", "АСД", "Скамбио Порте", "Лорд", "Аэлита", "Комфорт", "Двери регионов", "Термодур"):
        b = sorted(base_by_prod.get(p, []), key=lambda r: (r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", ""), r.get("image_url", "")))
        n = sorted(new_by_prod.get(p, []), key=lambda r: (r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", ""), r.get("image_url", "")))
        unchanged[p] = {"rows_before": len(b), "rows_after": len(n), "exact_match": b == n}

    decision_counts = Counter(r.get("decision", "") for r in review_rows)
    run_report = {
        "generated_at": datetime.now().isoformat(),
        "mode": "final_ready_v5_olymp_only",
        "source_catalog": str(BASE_CSV),
        "old_total_rows": len(base_rows),
        "total_rows": len(rows_sorted),
        "producers_count": len(producers),
        **coverage,
        "rows_with_image": sum(1 for r in rows_sorted if r.get("image_url", "")),
        "rows_without_image": sum(1 for r in rows_sorted if not r.get("image_url", "")),
        "needs_review_count": sum(1 for r in rows_sorted if r.get("status") == "needs_review"),
        "rows_by_producer": dict(sorted(Counter(r.get("producer", "") for r in rows_sorted if r.get("producer", "")).items())),
        "olymp_old_rows": len(olymp_old),
        "olymp_new_rows": len(olymp_new),
        "olymp_rows_with_image": sum(1 for r in olymp_new if r.get("image_url", "")),
        "olymp_rows_without_image": sum(1 for r in olymp_new if not r.get("image_url", "")),
        "olymp_photo_rows_needs_review": sum(1 for r in olymp_new if "— фото " in r.get("model_name", "")),
        "olymp_needs_review_count": sum(1 for r in olymp_new if r.get("status") == "needs_review"),
        "unmatched_review_counts": dict(decision_counts),
        "added_from_unmatched_or_new_scan": add_count,
        "rejected_candidates": reject_count,
        "duplicate_candidates": duplicate_count,
        "unchanged_producers_check": unchanged,
    }
    (OUT_DIR / "run_report.json").write_text(json.dumps(run_report, ensure_ascii=False, indent=2), encoding="utf-8")

    print("Built output/final_ready_v5")
    print(json.dumps({"olymp_old": len(olymp_old), "olymp_new": len(olymp_new), "with_image": run_report["olymp_rows_with_image"], "without_image": run_report["olymp_rows_without_image"]}, ensure_ascii=False))


if __name__ == "__main__":
    main()

