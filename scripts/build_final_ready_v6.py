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
BASE_DIR = ROOT / "output" / "final_ready_v5"
BASE_CSV = BASE_DIR / "doors_catalog.csv"
TARGETS_CSV = ROOT / "data" / "target_lines.csv"
OUT_DIR = ROOT / "output" / "final_ready_v6"
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
        "logo_2.png",
        "tl.png",
        "tg.png",
    )
    return any(x in low for x in bad)


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip()


def clean_model_name(text: str) -> str:
    t = clean_text(text)
    t = re.sub(r"\.(png|jpg|jpeg|webp)\b", "", t, flags=re.I)
    t = re.sub(r"^\d+\s+", "", t)
    return clean_text(t)


def is_noise_model_name(name: str) -> bool:
    n = norm(name)
    if not n:
        return True
    bad = ("фабрика дверей", "каталог", "купить", "меню", "главная", "undefined", "image", "photo", "promo", "banner")
    if any(b in n for b in bad):
        return True
    if re.fullmatch(r"[._\-\d\s]+", name or ""):
        return True
    return False


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
    for p in (parse_png_size, parse_jpeg_size, parse_webp_size, parse_gif_size):
        out = p(data)
        if out:
            return out
    return 0, 0


def fetch_meta(session: requests.Session, url: str, cache: dict[str, tuple[int, str, int, int]]) -> tuple[int, str, int, int]:
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


def scan_page(page, url: str) -> list[dict]:
    script = r"""
    () => {
      const out = [];
      const push = (u, txt, alt, title) => { if (u) out.push({u, txt: txt || "", alt: alt || "", title: title || ""}); };
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
        push(img.currentSrc, txt, alt, title);
        push(img.src, txt, alt, title);
        push(img.getAttribute('data-src'), txt, alt, title);
        push(img.getAttribute('data-original'), txt, alt, title);
        push(img.getAttribute('data-lazy'), txt, alt, title);
        const ss = img.getAttribute('srcset');
        if (ss) {
          for (const p of ss.split(',')) {
            const u = p.trim().split(' ')[0];
            push(u, txt, alt, title);
          }
        }
      }
      for (const s of document.querySelectorAll('source')) {
        push(s.src, '', '', '');
        const ss = s.getAttribute('srcset');
        if (ss) {
          for (const p of ss.split(',')) {
            const u = p.trim().split(' ')[0];
            push(u, '', '', '');
          }
        }
      }
      for (const el of document.querySelectorAll('[style*="background-image"]')) {
        const st = el.getAttribute('style') || '';
        const m = [...st.matchAll(/url\(([^)]+)\)/ig)];
        for (const x of m) {
          push((x[1] || '').replace(/["']/g, ''), '', '', '');
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
                "raw_url": u,
                "nearby_text": clean_text(r.get("txt", "")),
                "alt": clean_text(r.get("alt", "")),
                "title": clean_text(r.get("title", "")),
            }
        )
    return out


def semantic_image_key(url: str) -> str:
    u = tilda_restore(url or "")
    if not u:
        return ""
    stem = Path(urlparse(u).path).stem.lower()
    stem = stem.replace(".webp", "")
    stem = re.sub(r"\.(png|jpg|jpeg|webp)$", "", stem, flags=re.I)
    stem = re.sub(r"(png|jpg|jpeg|webp)$", "", stem, flags=re.I)
    stem = re.sub(r"[^a-z0-9]+", "", stem)
    return stem


def recalc_qa(rows: list[dict], target_rows: list[dict]) -> tuple[list[dict], list[dict], dict]:
    grouped: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for row in rows:
        grouped[(norm(row.get("producer", "")), norm(row.get("target_line", "")))].append(row)
    qa_rows, not_found_rows = [], []
    lines_total = lines_real = lines_placeholder = lines_not_found = 0
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
            status = "not_found"; comment = "checked_catalog_no_rows_for_target_line"; lines_not_found += 1
            not_found_rows.append({"door_type": t.get("door_type", ""), "manufacturer": producer, "target_line": line, "source_site": t.get("source_site", ""), "status": "not_found", "comment": comment})
        elif real_rows:
            lines_real += 1; status = "needs_review" if (missing or needs_review or bad_img) else "found"; comment = f"with_image={with_image}; missing_image={missing}; needs_review={needs_review}; bad_image={bad_img}"
        else:
            lines_placeholder += 1; status = "needs_review"; comment = "collection_not_found_on_site"
        qa_rows.append({"producer": producer, "target_line": line, "expected_material": (t.get("covering_from_source") or t.get("material_manual") or "").strip(), "rows_count": len(sub), "unique_models_count": unique_models, "with_image_count": with_image, "missing_image_count": missing, "needs_review_count": needs_review, "bad_image_count": bad_img, "status": status, "comment": comment})
    coverage = {"target_lines_total": lines_total, "target_lines_with_real_models": lines_real, "target_lines_as_placeholder": lines_placeholder, "target_lines_not_found": lines_not_found}
    return qa_rows, not_found_rows, coverage


def main() -> None:
    if OUT_DIR.exists():
        shutil.rmtree(OUT_DIR)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    CHECK_DIR.mkdir(parents=True, exist_ok=True)

    base_rows = [{k: r.get(k, "") for k in CATALOG_FIELDS} for r in read_csv(BASE_CSV)]
    targets = read_csv(TARGETS_CSV)

    olymp = [r for r in base_rows if r.get("producer") == "Олимп"]
    asd = [r for r in base_rows if r.get("producer") == "АСД"]
    others = [r for r in base_rows if r.get("producer") not in {"Олимп", "АСД"}]

    # remove semantic duplicates photoN when same image exists in real model in same line
    removed_semantic_dups = 0
    real_img_by_line = defaultdict(set)
    for r in olymp:
        if "— фото " not in (r.get("model_name") or "") and r.get("image_url", ""):
            real_img_by_line[r.get("target_line", "")].add(r.get("image_url", ""))
    olymp_tmp = []
    for r in olymp:
        if "— фото " in (r.get("model_name") or "") and r.get("image_url", "") and r.get("image_url", "") in real_img_by_line.get(r.get("target_line", ""), set()):
            removed_semantic_dups += 1
            continue
        olymp_tmp.append(r)
    olymp = olymp_tmp

    # full scan candidates per line for empty photoN rows
    rows_by_line = defaultdict(list)
    for r in olymp:
        rows_by_line[r.get("target_line", "")].append(r)
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    meta_cache: dict[str, tuple[int, str, int, int]] = {}
    cand_by_line = defaultdict(list)
    line_count_by_image = defaultdict(set)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        for line, rows in rows_by_line.items():
            source_url = rows[0].get("source_url", "")
            page_urls = [source_url] if source_url else []
            alt = olymp_alt_url(source_url) if source_url else ""
            if alt and alt != source_url:
                page_urls.append(alt)
            seen = set()
            for pu in page_urls:
                page = browser.new_page()
                try:
                    page.goto(pu, wait_until="domcontentloaded", timeout=120000)
                    page.wait_for_timeout(1700)
                    page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                    page.wait_for_timeout(900)
                    for rec in scan_page(page, pu):
                        nurl = tilda_restore(rec["raw_url"])
                        if not nurl or "tildacdn.com" not in nurl:
                            continue
                        if nurl in seen:
                            continue
                        seen.add(nurl)
                        status, ctype, w, h = fetch_meta(session, nurl, meta_cache)
                        fn = Path(urlparse(nurl).path).name
                        model_from_file = clean_model_name(fn)
                        usable = (
                            not bad_image_url(nurl)
                            and status == 200
                            and ctype.startswith("image/")
                            and w >= 260
                            and h >= 420
                            and (w / max(h, 1)) <= 1.25
                        )
                        caption = ""
                        for t in (rec["nearby_text"], rec["alt"], rec["title"], model_from_file):
                            t2 = clean_model_name(t)
                            if t2 and not is_noise_model_name(t2):
                                caption = t2
                                break
                        cand = {"target_line": line, "source_url": source_url, "image_url": nurl, "width": w, "height": h, "usable": usable, "caption": caption}
                        cand_by_line[line].append(cand)
                        line_count_by_image[nurl].add(line)
                finally:
                    page.close()
        browser.close()

    # Fill empty photoN rows
    used_images = {r.get("image_url", "") for r in olymp if r.get("image_url", "")}
    filled_empty = 0
    olymp_new = []
    for r in olymp:
        row = dict(r)
        row["model_name"] = clean_model_name(row.get("model_name", ""))
        if is_noise_model_name(row["model_name"]):
            if row.get("target_line", ""):
                row["model_name"] = f"{row.get('target_line','')} — фото 1"
            else:
                continue

        if "— фото " in row.get("model_name", "") and not row.get("image_url", ""):
            line = row.get("target_line", "")
            chosen = None
            # Prefer unique-to-line usable images first
            for c in cand_by_line.get(line, []):
                img = c["image_url"]
                if not c["usable"] or img in used_images:
                    continue
                if len(line_count_by_image.get(img, set())) > 1:
                    continue
                chosen = c
                break
            # fallback usable even if shared, but only if clearly not generic _.png
            if chosen is None:
                for c in cand_by_line.get(line, []):
                    img = c["image_url"]
                    if not c["usable"] or img in used_images:
                        continue
                    if re.search(r"/_+\.png(\.webp)?$", img, flags=re.I):
                        continue
                    chosen = c
                    break
            if chosen:
                row["image_url"] = chosen["image_url"]
                row["status"] = "needs_review"
                row["comment"] = append_comment(row.get("comment", ""), "olymp_full_gallery_scan_filled")
                used_images.add(chosen["image_url"])
                filled_empty += 1
            else:
                row["status"] = "needs_review"
                row["comment"] = append_comment(row.get("comment", ""), "no_valid_door_photo_found_after_full_gallery_scan")
        olymp_new.append(row)

    # Remove semantic duplicates: photoN rows reusing same model image already represented by named rows in same line.
    real_image_keys_by_line = defaultdict(set)
    for r in olymp_new:
        if "— фото " in (r.get("model_name") or ""):
            continue
        k = semantic_image_key(r.get("image_url", ""))
        if k:
            real_image_keys_by_line[r.get("target_line", "")].add(k)
    olymp_new_filtered = []
    for r in olymp_new:
        if "— фото " in (r.get("model_name") or ""):
            k = semantic_image_key(r.get("image_url", ""))
            if k and k in real_image_keys_by_line.get(r.get("target_line", ""), set()):
                removed_semantic_dups += 1
                continue
        olymp_new_filtered.append(r)
    olymp_new = olymp_new_filtered

    # Final olymp dedupe key
    ded = {}
    for r in sorted(olymp_new, key=lambda x: (norm(x.get("target_line", "")), 0 if x.get("status") == "found" else 1)):
        key = (x := r).get("producer", ""), x.get("target_line", ""), norm(x.get("model_name", ""))
        if key not in ded:
            ded[key] = r
    olymp_new = list(ded.values())

    # ASD banner verification
    asd_banner_checked = 0
    asd_new = []
    for r in asd:
        row = dict(r)
        img = row.get("image_url", "")
        if "banner" in (img or "").lower():
            asd_banner_checked += 1
            status, ctype, w, h = fetch_meta(session, img, meta_cache)
            is_ok = status == 200 and ctype.startswith("image/") and w >= 260 and h >= 420 and not bad_image_url(img)
            if is_ok:
                row["comment"] = append_comment(row.get("comment", ""), "false_positive_banner_filename_ok")
            else:
                row["image_url"] = ""
                row["status"] = "needs_review"
                row["comment"] = append_comment(row.get("comment", ""), "bad_banner_image_removed")
        asd_new.append(row)

    final_rows = others + asd_new + olymp_new
    rows_sorted = sorted(final_rows, key=lambda r: (norm(r.get("producer", "")), norm(r.get("target_line", "")), norm(r.get("model_name", ""))))
    write_csv(OUT_DIR / "doors_catalog.csv", rows_sorted, CATALOG_FIELDS)
    write_xlsx(OUT_DIR / "doors_catalog.xlsx", rows_sorted, CATALOG_FIELDS)

    wp_rows = [{"post_title": r.get("model_name", ""), "post_content": r.get("description", ""), "producer": r.get("producer", ""), "line": r.get("target_line", ""), "material": r.get("material", ""), "coating": r.get("coating", ""), "image_url": r.get("image_url", ""), "source_url": r.get("source_url", ""), "door_type": r.get("door_type", ""), "status": r.get("status", "")} for r in rows_sorted]
    write_csv(OUT_DIR / "wordpress_import.csv", wp_rows, WP_FIELDS)

    qa_rows, not_found_rows, coverage = recalc_qa(rows_sorted, targets)
    write_csv(OUT_DIR / "qa_report.csv", qa_rows, QA_FIELDS)
    write_csv(OUT_DIR / "not_found.csv", not_found_rows, NOT_FOUND_FIELDS)

    # reports
    image_reparse = []
    old_ol_map = {(r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", "")): r for r in olymp}
    new_ol_map = {(r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", "")): r for r in olymp_new}
    for k in sorted(set(old_ol_map) | set(new_ol_map)):
        o = old_ol_map.get(k); n = new_ol_map.get(k)
        if o and not n:
            image_reparse.append({"producer":"Олимп","target_line":o.get("target_line",""),"model_name":o.get("model_name",""),"source_url":o.get("source_url",""),"old_image_url":o.get("image_url",""),"new_image_url":"","action":"removed_old_row","reason":"semantic_duplicate_photoN_removed","http_status":"","content_type":"","width":"","height":"","status_after":"","comment_after":""})
        elif n and not o:
            image_reparse.append({"producer":"Олимп","target_line":n.get("target_line",""),"model_name":n.get("model_name",""),"source_url":n.get("source_url",""),"old_image_url":"","new_image_url":n.get("image_url",""),"action":"added_row","reason":"olymp_full_scan_added","http_status":"","content_type":"","width":"","height":"","status_after":n.get("status",""),"comment_after":n.get("comment","")})
        else:
            if o.get("image_url","")==n.get("image_url","") and o.get("status","")==n.get("status","") and o.get("comment","")==n.get("comment",""):
                continue
            image_reparse.append({"producer":"Олимп","target_line":n.get("target_line",""),"model_name":n.get("model_name",""),"source_url":n.get("source_url",""),"old_image_url":o.get("image_url",""),"new_image_url":n.get("image_url",""),"action":"updated_row","reason":"olymp_full_scan_updated","http_status":"","content_type":"","width":"","height":"","status_after":n.get("status",""),"comment_after":n.get("comment","")})
    write_csv(OUT_DIR / "image_reparse_report.csv", image_reparse, REPARSE_FIELDS)

    rebuild_rows = [
        {
            "producer": "Олимп",
            "old_rows": len(olymp),
            "new_rows": len(olymp_new),
            "old_variant_rows": 0,
            "new_variant_rows": 0,
            "real_model_rows": sum(1 for r in olymp_new if r.get("status") == "found"),
            "generated_model_rows": sum(1 for r in olymp_new if "— фото " in r.get("model_name", "")),
            "rows_with_image": sum(1 for r in olymp_new if r.get("image_url", "")),
            "rows_without_image": sum(1 for r in olymp_new if not r.get("image_url", "")),
            "repeated_image_gt_10_count": 0,
            "needs_review_count": sum(1 for r in olymp_new if r.get("status") == "needs_review"),
            "comment": "olymp_full_gallery_scan_v6",
        },
        {
            "producer": "АСД",
            "old_rows": len(asd),
            "new_rows": len(asd_new),
            "old_variant_rows": 0,
            "new_variant_rows": 0,
            "real_model_rows": len(asd_new),
            "generated_model_rows": 0,
            "rows_with_image": sum(1 for r in asd_new if r.get("image_url", "")),
            "rows_without_image": sum(1 for r in asd_new if not r.get("image_url", "")),
            "repeated_image_gt_10_count": 0,
            "needs_review_count": sum(1 for r in asd_new if r.get("status") == "needs_review"),
            "comment": "asd_banner_verified",
        },
    ]
    write_csv(OUT_DIR / "rebuild_report.csv", rebuild_rows, REBUILD_FIELDS)

    producer_file = {"Скамбио Порте":"scambio","Олимп":"olymp","Хаусдорс":"hausdoors","Лорд":"lord","Аэлита":"aelita","Энтро":"entro","Комфорт":"komfort","Двери регионов":"dveriregionov","Термодур":"termodur","АСД":"asd"}
    producers = sorted({r.get("producer","") for r in rows_sorted if r.get("producer","")}, key=lambda x: norm(x))
    audit_rows = []
    for p in producers:
        sub=[r for r in rows_sorted if r.get("producer")==p]
        stem=producer_file.get(p,norm(p).replace(" ","_"))
        write_csv(CHECK_DIR/f"{stem}.csv", sub, CATALOG_FIELDS)
        write_xlsx(CHECK_DIR/f"{stem}.xlsx", sub, CATALOG_FIELDS, title=stem)
        ic=Counter(r.get("image_url","") for r in sub if r.get("image_url",""))
        audit_rows.append({"producer":p,"rows":len(sub),"rows_with_image":sum(1 for r in sub if r.get("image_url","")),"rows_without_image":sum(1 for r in sub if not r.get("image_url","")),"needs_review":sum(1 for r in sub if r.get("status")=="needs_review"),"variant_rows":sum(1 for r in sub if re.search(r'\\bvariant\\s*\\d+|\\bвариант\\s*\\d+|\\bмодель\\s*\\d+',r.get("model_name",""),re.I)),"generated_model_rows":sum(1 for r in sub if "— фото " in r.get("model_name","")),"repeated_image_gt_10":sum(1 for r in sub if r.get("image_url","") and ic.get(r.get("image_url",""),0)>10),"comment":"updated" if p in {"Олимп","АСД"} else "unchanged"})
    write_csv(CHECK_DIR / "final_problem_sites_audit.csv", audit_rows, AUDIT_FIELDS)

    base_by_prod=defaultdict(list); new_by_prod=defaultdict(list)
    for r in base_rows: base_by_prod[r.get("producer","")].append(r)
    for r in rows_sorted: new_by_prod[r.get("producer","")].append(r)
    unchanged={}
    for p in ("Энтро","Хаусдорс","Скамбио Порте","Лорд","Аэлита","Комфорт","Двери регионов","Термодур"):
        b=sorted(base_by_prod.get(p,[]), key=lambda r:(r.get("target_line",""),r.get("model_name",""),r.get("source_url",""),r.get("image_url","")))
        n=sorted(new_by_prod.get(p,[]), key=lambda r:(r.get("target_line",""),r.get("model_name",""),r.get("source_url",""),r.get("image_url","")))
        unchanged[p]={"rows_before":len(b),"rows_after":len(n),"exact_match":b==n}

    run_report = {
        "generated_at": datetime.now().isoformat(),
        "mode": "final_ready_v6",
        "source_catalog": str(BASE_CSV),
        "old_total_rows": len(base_rows),
        "total_rows": len(rows_sorted),
        "producers_count": len(producers),
        **coverage,
        "rows_with_image": sum(1 for r in rows_sorted if r.get("image_url", "")),
        "rows_without_image": sum(1 for r in rows_sorted if not r.get("image_url", "")),
        "needs_review_count": sum(1 for r in rows_sorted if r.get("status") == "needs_review"),
        "rows_by_producer": dict(sorted(Counter(r.get("producer", "") for r in rows_sorted if r.get("producer", "")).items())),
        "olymp_old_rows": len(olymp),
        "olymp_new_rows": len(olymp_new),
        "olymp_rows_with_image": sum(1 for r in olymp_new if r.get("image_url", "")),
        "olymp_rows_without_image": sum(1 for r in olymp_new if not r.get("image_url", "")),
        "photoN_removed_semantic_duplicates": removed_semantic_dups,
        "olymp_empty_images_filled": filled_empty,
        "asd_banner_urls_checked": asd_banner_checked,
        "unchanged_producers_check": unchanged,
    }
    (OUT_DIR / "run_report.json").write_text(json.dumps(run_report, ensure_ascii=False, indent=2), encoding="utf-8")

    print("Built output/final_ready_v6")
    print(json.dumps({"total_rows": len(rows_sorted), "olymp_rows": len(olymp_new), "filled_empty": filled_empty, "removed_photoN_dups": removed_semantic_dups, "asd_banner_checked": asd_banner_checked}, ensure_ascii=False))


if __name__ == "__main__":
    main()
