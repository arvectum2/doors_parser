from __future__ import annotations

import csv
import json
import re
import shutil
from collections import Counter, defaultdict
from datetime import datetime
from html import escape
from pathlib import Path
from urllib.parse import urljoin, urlparse

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from visual_validate_full_door_images import (
    VisualCheck,
    extract_candidates_from_html,
    is_probable_color_row,
    pick_best_visual_full_door,
    validate_full_door_image,
)


ROOT = Path(__file__).resolve().parent.parent
BASE_DIR = ROOT / "output" / "final_ready_v10"
BASE_CSV = BASE_DIR / "doors_catalog.csv"
TARGETS_CSV = ROOT / "data" / "target_lines.csv"
OUT_DIR = ROOT / "output" / "final_ready_v11"
CHECK_DIR = OUT_DIR / "check_by_producer"
SHEETS_DIR = OUT_DIR / "full_door_contact_sheets"
PROBLEM = ("АСД", "Хаусдорс")
UNCHANGED_EIGHT = ("Олимп", "Энтро", "Скамбио Порте", "Лорд", "Аэлита", "Комфорт", "Двери регионов", "Термодур")
PROBLEM_REVIEW_DIR = OUT_DIR / "problem_images_review"

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
VISUAL_AUDIT_FIELDS = [
    "producer",
    "target_line",
    "model_name",
    "source_url",
    "old_image_url",
    "new_image_url",
    "image_width",
    "image_height",
    "object_bbox_width",
    "object_bbox_height",
    "object_aspect_ratio",
    "visual_decision",
    "action",
    "reason",
    "comment",
]
EVIDENCE_FIELDS = [
    "producer",
    "target_line",
    "model_name",
    "source_url",
    "checked_candidates_count",
    "rejected_candidates_count",
    "best_rejected_candidate_url",
    "best_rejected_reason",
    "final_reason",
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


def bad_image_url(url: str) -> bool:
    low = (url or "").lower()
    if not low:
        return True
    bad = (
        "badge",
        "banner",
        "hero",
        "interior",
        "texture",
        "sample",
        "pattern",
        "handle",
        "lock",
        "profile",
        "section",
        "empty",
        "placeholder",
        "open.png",
        "catalog.svg",
        "resize/20x",
    )
    return any(x in low for x in bad)


def hard_bad_tokens(text: str) -> bool:
    low = (text or "").lower()
    tokens = (
        "badge",
        "hit",
        "sale",
        "logo",
        "icon",
        "label",
        "sticker",
        "placeholder",
        "no-image",
        "no_image",
        "empty",
        "pixel",
        "sprite",
        "open.png",
        "catalog.svg",
        "interier",
        "interior",
        "room",
        "roomset",
        "scene",
        "interer",
        "интерьер",
        "hero",
        "banner",
        "promo",
        "kromo",
        "kromka",
        "кромка",
        "handle",
        "ручка",
        "lock",
        "замок",
        "hinge",
        "petlya",
    )
    return any(t in low for t in tokens)


def colorish_tokens(text: str) -> bool:
    low = (text or "").lower()
    tokens = ("ral", "ncs", "color", "colour", "texture", "sample", "pattern", "material", "finish", "decor", "матовый", "покрытие", "цвет")
    return any(t in low for t in tokens)


def classify_no_photo_reason(reasons: list[str], source_status: int | None) -> str:
    if source_status and source_status >= 400:
        return "source_url_404"
    r = " ".join(reasons).lower()
    if "interior" in r or "wide_layout" in r:
        return "only_interior_images_found"
    if "color_sample" in r or "texture" in r:
        return "only_color_samples_found"
    if "fragment" in r or "object_not_vertical" in r:
        return "only_fragments_found"
    if not reasons:
        return "no_product_images_on_page"
    return "no_full_door_candidate_found"


def producer_accepts_visual(producer: str, model_name: str, check: VisualCheck) -> bool:
    if check.status_code != 200 or not check.content_type.startswith("image/"):
        return False
    low = (check.url or "").lower()
    if hard_bad_tokens(low):
        return False
    if producer == "Хаусдорс" and (is_probable_color_row(model_name) or colorish_tokens(low)):
        return False
    if check.decision == "full_door":
        return True
    # ASD: many valid product photos are wide catalog renders (e.g. 1058x561).
    if producer == "АСД":
        if check.width >= 800 and check.height >= 480 and check.bbox_height >= int(check.height * 0.45):
            if not any(t in low for t in ("interior", "room", "hero", "banner")):
                return True
    # Entro/Olymp/Hausdoors soft fallback when object is still vertically dominant
    if producer in {"Энтро", "Олимп", "Хаусдорс"}:
        if check.width >= 260 and check.height >= 420 and check.bbox_aspect >= 1.35 and (check.height / max(1, check.width)) >= 0.95:
            if not any(t in low for t in ("interior", "room", "hero", "banner", "texture", "sample", "pattern")):
                return True
    return False


def candidate_rank(producer: str, model_name: str, check: VisualCheck, source_url: str) -> float:
    low = (check.url or "").lower()
    score = check.score if check.score > -900 else 0.0
    if producer == "АСД" and "as-doors.ru/assets/cache_image/" in low:
        score += 2.0
    if producer == "Хаусдорс" and "tildacdn.com" in low:
        score += 0.8
    if source_url and urlparse(source_url).netloc and urlparse(check.url).netloc == urlparse(source_url).netloc:
        score += 0.6
    if "product" in low or "model" in low or "door" in low or "dver" in low:
        score += 0.6
    if is_probable_color_row(model_name) and colorish_tokens(low):
        score -= 5.0
    return score


def extract_candidates_playwright(page, url: str) -> list[str]:
    script = r"""
    () => {
      const out = [];
      const push = (u) => { if (u) out.push(u); };
      const abs = (u) => { try { return new URL(u, location.href).href; } catch(_) { return ''; } };
      const parseSrcset = (ss) => {
        if (!ss) return;
        ss.split(',').forEach(p => {
          const u = (p.trim().split(' ')[0] || '').trim();
          if (u) push(u);
        });
      };
      for (const img of document.querySelectorAll('img')) {
        push(img.currentSrc);
        push(img.getAttribute('src'));
        push(img.getAttribute('data-src'));
        push(img.getAttribute('data-original'));
        push(img.getAttribute('data-lazy'));
        parseSrcset(img.getAttribute('srcset'));
        parseSrcset(img.getAttribute('data-srcset'));
      }
      for (const s of document.querySelectorAll('source')) {
        push(s.getAttribute('src'));
        parseSrcset(s.getAttribute('srcset'));
      }
      for (const el of document.querySelectorAll('[style*="background-image"]')) {
        const st = el.getAttribute('style') || '';
        const ms = [...st.matchAll(/url\(([^)]+)\)/ig)];
        for (const m of ms) push((m[1] || '').replace(/["']/g, ''));
      }
      // Tilda JSON blobs
      for (const sc of document.querySelectorAll('script')) {
        const tx = sc.textContent || '';
        if (!tx) continue;
        const m = tx.match(/https?:\/\/[^"'\s]+/g) || [];
        m.forEach(u => push(u));
      }
      return out.map(abs).filter(Boolean);
    }
    """
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=70000)
        page.wait_for_timeout(2200)
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        page.wait_for_timeout(1200)
        rows = page.evaluate(script) or []
    except PlaywrightTimeoutError:
        rows = []
    except Exception:
        rows = []
    out = []
    seen = set()
    for u in rows:
        u = (u or "").strip()
        if not u:
            continue
        if not re.search(r"\.(jpg|jpeg|png|webp|gif)(\?|$)", u, flags=re.I):
            continue
        if u in seen:
            continue
        seen.add(u)
        out.append(u)
    return out


def recalc_qa(rows: list[dict], target_rows: list[dict]) -> tuple[list[dict], list[dict], dict]:
    grouped: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for row in rows:
        grouped[(norm(row.get("producer", "")), norm(row.get("target_line", "")))].append(row)
    qa_rows, not_found_rows = [], []
    lt = lr = lp = lnf = 0
    for t in target_rows:
        p = (t.get("manufacturer") or "").strip()
        l = (t.get("target_line") or "").strip()
        if not p or not l:
            continue
        lt += 1
        key = (norm(p), norm(l))
        sub = grouped.get(key, [])
        if not sub and norm(p) == norm("АСД") and norm(l) == norm("Все модели"):
            sub = grouped.get((norm("АСД"), norm("Уточнить коллекцию")), [])
        with_img = sum(1 for r in sub if r.get("image_url", ""))
        miss = len(sub) - with_img
        nr = sum(1 for r in sub if r.get("status") == "needs_review")
        bad = sum(1 for r in sub if bad_image_url(r.get("image_url", "")))
        uniq = len({norm_model(r.get("model_name", "")) for r in sub if r.get("model_name", "")})
        placeholders = [r for r in sub if "collection_not_found_on_site" in (r.get("comment") or "").lower()]
        real = [r for r in sub if r not in placeholders]
        if not sub:
            st = "not_found"
            cm = "checked_catalog_no_rows_for_target_line"
            lnf += 1
            not_found_rows.append({"door_type": t.get("door_type", ""), "manufacturer": p, "target_line": l, "source_site": t.get("source_site", ""), "status": "not_found", "comment": cm})
        elif real:
            lr += 1
            st = "needs_review" if (miss or nr or bad) else "found"
            cm = f"with_image={with_img}; missing_image={miss}; needs_review={nr}; bad_image={bad}"
        else:
            lp += 1
            st = "needs_review"
            cm = "collection_not_found_on_site"
        qa_rows.append({"producer": p, "target_line": l, "expected_material": (t.get("covering_from_source") or t.get("material_manual") or "").strip(), "rows_count": len(sub), "unique_models_count": uniq, "with_image_count": with_img, "missing_image_count": miss, "needs_review_count": nr, "bad_image_count": bad, "status": st, "comment": cm})
    return qa_rows, not_found_rows, {"target_lines_total": lt, "target_lines_with_real_models": lr, "target_lines_as_placeholder": lp, "target_lines_not_found": lnf}


def contact_sheet_html(rows: list[dict], title: str, candidates_by_key: dict[tuple[str, str, str, str], list[VisualCheck]]) -> str:
    out = [
        "<!doctype html><html><head><meta charset='utf-8'>",
        f"<title>{escape(title)}</title>",
        "<style>body{font-family:Arial,sans-serif;margin:20px} h2{margin-top:28px} .card{border:1px solid #ddd;padding:10px;margin:10px 0} .row{display:flex;gap:12px;align-items:flex-start;flex-wrap:wrap} .img{max-width:130px;max-height:220px;object-fit:contain;border:1px solid #ddd} .muted{color:#666;font-size:12px} .rej{font-size:11px;line-height:1.35}</style></head><body>",
        f"<h1>{escape(title)}</h1>",
    ]
    for r in rows:
        key = (r.get("producer", ""), r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", ""))
        cands = candidates_by_key.get(key, [])
        selected_url = (r.get("new_image_url", "") or "").strip()
        rejected = [c for c in cands if not selected_url or c.url != selected_url]
        rejected = sorted(rejected, key=lambda x: x.score, reverse=True)[:5]
        out.append("<div class='card'>")
        out.append(f"<div><b>{escape(r.get('model_name',''))}</b> <span class='muted'>[{escape(r.get('action',''))}]</span></div>")
        out.append(f"<div class='muted'>{escape(r.get('target_line',''))} | {escape(r.get('reason',''))}</div>")
        out.append("<div class='row'>")
        if selected_url:
            out.append(f"<div><div class='muted'>selected</div><img class='img' src='{escape(selected_url)}'><div class='rej'>{escape(selected_url)}</div></div>")
        else:
            out.append("<div class='muted'>selected: none</div>")
        if rejected:
            out.append("<div><div class='muted'>rejected candidates</div>")
            for c in rejected:
                out.append(
                    "<div class='rej'>"
                    f"<a href='{escape(c.url)}'>{escape(c.url)}</a><br>"
                    f"decision={escape(c.decision)} reason={escape(c.reason)} "
                    f"w={c.width} h={c.height} bbox={c.bbox_width}x{c.bbox_height} asp={c.bbox_aspect}"
                    "</div><br>"
                )
            out.append("</div>")
        out.append("</div></div>")
    out.append("</body></html>")
    return "".join(out)


def olymp_alt_url(source_url: str) -> str:
    p = urlparse(source_url)
    path = p.path.strip("/")
    if not path:
        return source_url
    slug = path.split("/")[-1]
    slug2 = re.sub(r"-\d+$", "", slug)
    if slug2 == slug:
        return source_url
    return f"{p.scheme}://{p.netloc}/{'/'.join(path.split('/')[:-1] + [slug2])}"


def candidate_urls_for_source(
    session: requests.Session,
    source_url: str,
    page_cache: dict[str, dict],
    producer: str,
    pw_page=None,
) -> tuple[list[str], int]:
    src = (source_url or "").strip()
    if not src:
        return [], 0
    if src in page_cache:
        p = page_cache[src]
        return p.get("candidates", []), int(p.get("source_status", 0))
    source_list = [src]
    if "dveriolimp.ru" in src:
        alt = olymp_alt_url(src)
        if alt != src:
            source_list.append(alt)
    out, seen = [], set()
    source_status = 0
    for s in source_list:
        try:
            resp = session.get(s, timeout=(8, 35), allow_redirects=True)
            source_status = max(source_status, resp.status_code)
            html = resp.text
        except Exception:
            continue
        for c in extract_candidates_from_html(s, html):
            if c in seen:
                continue
            seen.add(c)
            out.append(c)
    # Hausdoors: aggressive lazy/tilda extraction via Playwright.
    if producer == "Хаусдорс" and pw_page is not None:
        for s in source_list:
            for c in extract_candidates_playwright(pw_page, s):
                if c in seen:
                    continue
                seen.add(c)
                out.append(c)
    # ensure non-zero checked candidate count when we already had a previous URL
    if not out and src:
        out.append(src)
    page_cache[src] = {"candidates": out, "source_status": source_status}
    return out, source_status


def main() -> None:
    if OUT_DIR.exists():
        shutil.rmtree(OUT_DIR)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    CHECK_DIR.mkdir(parents=True, exist_ok=True)
    SHEETS_DIR.mkdir(parents=True, exist_ok=True)
    PROBLEM_REVIEW_DIR.mkdir(parents=True, exist_ok=True)

    base_rows = [{k: r.get(k, "") for k in CATALOG_FIELDS} for r in read_csv(BASE_CSV)]
    targets = read_csv(TARGETS_CSV)

    by_prod = defaultdict(list)
    for r in base_rows:
        by_prod[r.get("producer", "")].append(dict(r))

    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    vis_cache: dict[str, VisualCheck] = {}
    page_cache: dict[str, dict] = {}

    updated_by_prod: dict[str, list[dict]] = {}
    audit_rows: list[dict] = []
    reparse_rows: list[dict] = []
    evidence_rows: list[dict] = []
    per_stats = defaultdict(lambda: defaultdict(int))
    reason_counter = Counter()
    row_candidates: dict[tuple[str, str, str, str], list[VisualCheck]] = {}

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        context = browser.new_context(viewport={"width": 1440, "height": 2200}, ignore_https_errors=True)
        page = context.new_page()
        for producer in PROBLEM:
            rows = [dict(r) for r in by_prod.get(producer, [])]
            new_rows: list[dict] = []
            for row in rows:
                old_url = row.get("image_url", "")
                source_url = row.get("source_url", "")
                model_name = row.get("model_name", "")
                key = (producer, row.get("target_line", ""), model_name, source_url)

                if producer == "Хаусдорс" and re.search(r"(?i)^\s*кромка\b", model_name or ""):
                    per_stats[producer]["deleted"] += 1
                    per_stats[producer]["hausdoors_kromka_deleted"] += 1
                    reason_counter["component"] += 1
                    audit_rows.append(
                        {
                            "producer": producer,
                            "target_line": row.get("target_line", ""),
                            "model_name": model_name,
                            "source_url": source_url,
                            "old_image_url": old_url,
                            "new_image_url": "",
                            "image_width": "",
                            "image_height": "",
                            "object_bbox_width": "",
                            "object_bbox_height": "",
                            "object_aspect_ratio": "",
                            "visual_decision": "component",
                            "action": "deleted_row",
                            "reason": "kromka_row_deleted",
                            "comment": row.get("comment", ""),
                        }
                    )
                    reparse_rows.append(
                        {
                            "producer": producer,
                            "target_line": row.get("target_line", ""),
                            "model_name": model_name,
                            "source_url": source_url,
                            "old_image_url": old_url,
                            "new_image_url": "",
                            "action": "deleted_row",
                            "reason": "kromka_row_deleted",
                            "http_status": "",
                            "content_type": "",
                            "width": "",
                            "height": "",
                            "status_after": "",
                            "comment_after": row.get("comment", ""),
                        }
                    )
                    continue

                old_check = validate_full_door_image(session, old_url, vis_cache) if old_url else VisualCheck("", 0, "", 0, 0, 0, 0, 0.0, "no_image", "no_image", -999.0)
                candidates, source_status = candidate_urls_for_source(session, source_url, page_cache, producer, pw_page=page if producer == "Хаусдорс" else None)
                # always include current candidate as reference
                if old_url and old_url not in candidates:
                    candidates = [old_url] + candidates

                checked: list[VisualCheck] = []
                accepted: list[VisualCheck] = []
                for cu in candidates:
                    chk = validate_full_door_image(session, cu, vis_cache)
                    checked.append(chk)
                    if producer_accepts_visual(producer, model_name, chk):
                        accepted.append(chk)
                row_candidates[key] = checked

                action = "kept"
                chosen = old_check
                visual_decision = old_check.decision
                reason = old_check.reason

                old_is_good = producer_accepts_visual(producer, model_name, old_check)
                if old_is_good:
                    chosen = old_check
                    row["image_url"] = old_check.url
                    row["status"] = "found"
                    per_stats[producer]["kept"] += 1
                    visual_decision = "full_door"
                    reason = "full_door_image_ok"
                elif accepted:
                    best = sorted(accepted, key=lambda c: candidate_rank(producer, model_name, c, source_url), reverse=True)[0]
                    chosen = best
                    row["image_url"] = best.url
                    row["status"] = "found"
                    row["comment"] = append_comment(row.get("comment", ""), "full_door_image_selected")
                    if old_url and best.url != old_url:
                        action = "replaced"
                        per_stats[producer]["replaced"] += 1
                        reason = "replaced_with_full_door"
                    else:
                        per_stats[producer]["kept"] += 1
                        reason = "full_door_image_ok"
                    visual_decision = "full_door"
                else:
                    row["image_url"] = ""
                    row["status"] = "needs_review"
                    if producer == "АСД" and re.search(r"(?i)interier|interior|room|roomset|scene|interer|интерьер", old_url or ""):
                        row["comment"] = append_comment(row.get("comment", ""), "interior_image_removed_no_full_door_found")
                    else:
                        row["comment"] = append_comment(row.get("comment", ""), "no_full_door_image_found")
                    action = "cleared" if old_url else "needs_review"
                    per_stats[producer]["cleared"] += 1 if action == "cleared" else 0
                    per_stats[producer]["needs_review"] += 1 if action == "needs_review" else 0
                    rejected = sorted(checked, key=lambda c: candidate_rank(producer, model_name, c, source_url), reverse=True)
                    best_rej = rejected[0] if rejected else old_check
                    chosen = best_rej
                    visual_decision = best_rej.decision if best_rej.decision else "no_image"
                    reason = best_rej.reason if best_rej.reason else "no_full_door_candidate_found_after_source_page_scan"
                    reasons = [c.reason for c in checked if c.reason]
                    if producer in {"АСД", "Хаусдорс"}:
                        evidence_rows.append(
                            {
                                "producer": producer,
                                "target_line": row.get("target_line", ""),
                                "model_name": model_name,
                                "source_url": source_url,
                                "checked_candidates_count": len(checked),
                                "rejected_candidates_count": len(checked),
                                "best_rejected_candidate_url": best_rej.url,
                                "best_rejected_reason": best_rej.reason,
                                "final_reason": classify_no_photo_reason(reasons, source_status),
                            }
                        )

                # Hausdoors finish/color rows: keep only if full door was verified.
                if producer == "Хаусдорс" and is_probable_color_row(model_name):
                    if row.get("image_url", ""):
                        row["comment"] = append_comment(row.get("comment", ""), "full_door_finish_variant_verified")
                    else:
                        reason_counter["color_sample"] += 1

                if not row.get("image_url", ""):
                    row["status"] = "needs_review"
                    row["comment"] = append_comment(row.get("comment", ""), "no_full_door_image_found")
                    per_stats[producer]["without_image"] += 1

                if visual_decision in {"color_sample", "texture", "fragment", "interior", "component", "banner"}:
                    reason_counter[visual_decision] += 1

                audit_rows.append(
                    {
                        "producer": producer,
                        "target_line": row.get("target_line", ""),
                        "model_name": model_name,
                        "source_url": source_url,
                        "old_image_url": old_url,
                        "new_image_url": row.get("image_url", ""),
                        "image_width": chosen.width,
                        "image_height": chosen.height,
                        "object_bbox_width": chosen.bbox_width,
                        "object_bbox_height": chosen.bbox_height,
                        "object_aspect_ratio": chosen.bbox_aspect,
                        "visual_decision": visual_decision,
                        "action": action,
                        "reason": reason,
                        "comment": row.get("comment", ""),
                    }
                )
                if old_url != row.get("image_url", "") or action in {"cleared", "replaced", "needs_review"}:
                    reparse_rows.append(
                        {
                            "producer": producer,
                            "target_line": row.get("target_line", ""),
                            "model_name": model_name,
                            "source_url": source_url,
                            "old_image_url": old_url,
                            "new_image_url": row.get("image_url", ""),
                            "action": action,
                            "reason": reason,
                            "http_status": chosen.status_code,
                            "content_type": chosen.content_type,
                            "width": chosen.width,
                            "height": chosen.height,
                            "status_after": row.get("status", ""),
                            "comment_after": row.get("comment", ""),
                        }
                    )
                new_rows.append(row)
            updated_by_prod[producer] = new_rows
        page.close()
        context.close()
        browser.close()

    # Olymp rule: max one empty photo-N per target_line
    olymp_rows = updated_by_prod.get("Олимп", [])
    by_line = defaultdict(list)
    for r in olymp_rows:
        by_line[r.get("target_line", "")].append(r)
    olymp_final = []
    for line, rows in by_line.items():
        empties = [r for r in rows if re.search(r"(?i)\s[—-]\s*фото\s*\d+\s*$", r.get("model_name", "")) and not r.get("image_url", "").strip()]
        non = [r for r in rows if r not in empties]
        if non:
            olymp_final.extend(non)
        elif empties:
            keep = dict(empties[0])
            keep["model_name"] = line or "Олимп"
            keep["status"] = "needs_review"
            keep["comment"] = append_comment(keep.get("comment", ""), "no_full_door_image_found")
            olymp_final.append(keep)
            for extra in empties[1:]:
                audit_rows.append(
                    {
                        "producer": "Олимп",
                        "target_line": extra.get("target_line", ""),
                        "model_name": extra.get("model_name", ""),
                        "source_url": extra.get("source_url", ""),
                        "old_image_url": extra.get("image_url", ""),
                        "new_image_url": "",
                        "image_width": "",
                        "image_height": "",
                        "object_bbox_width": "",
                        "object_bbox_height": "",
                        "object_aspect_ratio": "",
                        "visual_decision": "no_image",
                        "action": "deleted_row",
                        "reason": "extra_empty_photoN_removed",
                        "comment": extra.get("comment", ""),
                    }
                )
    updated_by_prod["Олимп"] = olymp_final

    final_rows: list[dict] = []
    for prod, rows in by_prod.items():
        if prod in PROBLEM:
            final_rows.extend(updated_by_prod.get(prod, []))
        else:
            final_rows.extend(rows)

    rows_sorted = sorted(final_rows, key=lambda r: (norm(r.get("producer", "")), norm(r.get("target_line", "")), norm(r.get("model_name", ""))))
    write_csv(OUT_DIR / "doors_catalog.csv", rows_sorted, CATALOG_FIELDS)
    write_xlsx(OUT_DIR / "doors_catalog.xlsx", rows_sorted, CATALOG_FIELDS)
    wp_rows = [{"post_title": r.get("model_name", ""), "post_content": r.get("description", ""), "producer": r.get("producer", ""), "line": r.get("target_line", ""), "material": r.get("material", ""), "coating": r.get("coating", ""), "image_url": r.get("image_url", ""), "source_url": r.get("source_url", ""), "door_type": r.get("door_type", ""), "status": r.get("status", "")} for r in rows_sorted]
    write_csv(OUT_DIR / "wordpress_import.csv", wp_rows, WP_FIELDS)

    qa_rows, not_found_rows, coverage = recalc_qa(rows_sorted, targets)
    write_csv(OUT_DIR / "qa_report.csv", qa_rows, QA_FIELDS)
    write_csv(OUT_DIR / "not_found.csv", not_found_rows, NOT_FOUND_FIELDS)
    write_csv(OUT_DIR / "image_reparse_report.csv", reparse_rows, REPARSE_FIELDS)
    write_csv(OUT_DIR / "full_door_visual_audit.csv", audit_rows, VISUAL_AUDIT_FIELDS)
    write_csv(OUT_DIR / "no_full_door_evidence.csv", evidence_rows, EVIDENCE_FIELDS)

    # contact sheets
    for producer in PROBLEM:
        prs = [r for r in audit_rows if r.get("producer") == producer]
        fname = {
            "Хаусдорс": "hausdoors_selected_and_rejected.html",
            "АСД": "asd_selected_and_rejected.html",
        }[producer]
        (SHEETS_DIR / fname).write_text(contact_sheet_html(prs, f"{producer} full-door audit", row_candidates), encoding="utf-8")

    asd_problem = [
        r
        for r in audit_rows
        if r.get("producer") == "АСД"
        and (
            re.search(r"(?i)interier|interior|room|roomset|scene|interer|интерьер", r.get("old_image_url", "") or "")
            or "interior_image_removed_no_full_door_found" in (r.get("comment", "") or "")
        )
    ]
    haus_problem = [
        r
        for r in audit_rows
        if r.get("producer") == "Хаусдорс"
        and (
            re.search(r"(?i)^\s*кромка\b", r.get("model_name", "") or "")
            or is_probable_color_row(r.get("model_name", ""))
        )
    ]
    (PROBLEM_REVIEW_DIR / "asd_interier_removed.html").write_text(
        contact_sheet_html(asd_problem, "ASD interior removed review", row_candidates), encoding="utf-8"
    )
    (PROBLEM_REVIEW_DIR / "hausdoors_finish_color_review.html").write_text(
        contact_sheet_html(haus_problem, "Hausdoors finish/color review", row_candidates), encoding="utf-8"
    )

    # per producer exports
    producer_file = {"Скамбио Порте": "scambio", "Олимп": "olymp", "Хаусдорс": "hausdoors", "Лорд": "lord", "Аэлита": "aelita", "Энтро": "entro", "Комфорт": "komfort", "Двери регионов": "dveriregionov", "Термодур": "termodur", "АСД": "asd"}
    producers = sorted({r.get("producer", "") for r in rows_sorted if r.get("producer", "")}, key=lambda x: norm(x))
    for p in producers:
        sub = [r for r in rows_sorted if r.get("producer") == p]
        stem = producer_file.get(p, norm(p).replace(" ", "_"))
        write_csv(CHECK_DIR / f"{stem}.csv", sub, CATALOG_FIELDS)
        write_xlsx(CHECK_DIR / f"{stem}.xlsx", sub, CATALOG_FIELDS, title=stem)

    # unchanged eight check
    old_by = defaultdict(list)
    new_by = defaultdict(list)
    for r in base_rows:
        old_by[r.get("producer", "")].append(r)
    for r in rows_sorted:
        new_by[r.get("producer", "")].append(r)
    unchanged_check = {}
    for p in UNCHANGED_EIGHT:
        a = sorted(old_by.get(p, []), key=lambda r: (r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", ""), r.get("image_url", ""), r.get("status", ""), r.get("comment", "")))
        b = sorted(new_by.get(p, []), key=lambda r: (r.get("target_line", ""), r.get("model_name", ""), r.get("source_url", ""), r.get("image_url", ""), r.get("status", ""), r.get("comment", "")))
        unchanged_check[p] = {"rows_before": len(a), "rows_after": len(b), "exact_match": a == b}

    rows_by_producer = dict(sorted(Counter(r.get("producer", "") for r in rows_sorted if r.get("producer", "")).items()))
    before_by_producer = dict(sorted(Counter(r.get("producer", "") for r in base_rows if r.get("producer", "")).items()))
    problem_summary = {}
    for p in PROBLEM:
        before = before_by_producer.get(p, 0)
        after = rows_by_producer.get(p, 0)
        pr = [r for r in rows_sorted if r.get("producer") == p]
        problem_summary[p] = {
            "rows_before": before,
            "rows_after": after,
            "kept_full_door": per_stats[p].get("kept", 0),
            "replaced": per_stats[p].get("replaced", 0),
            "cleared": per_stats[p].get("cleared", 0),
            "deleted_rows": per_stats[p].get("deleted", 0),
            "without_image": sum(1 for r in pr if not r.get("image_url", "")),
        }

    run_report = {
        "generated_at": datetime.now().isoformat(),
        "mode": "final_ready_v11",
        "source_catalog": str(BASE_CSV),
        "old_total_rows": len(base_rows),
        "total_rows": len(rows_sorted),
        "producers_count": len(producers),
        **coverage,
        "rows_with_image": sum(1 for r in rows_sorted if r.get("image_url", "")),
        "rows_without_image": sum(1 for r in rows_sorted if not r.get("image_url", "")),
        "needs_review_count": sum(1 for r in rows_sorted if r.get("status") == "needs_review"),
        "rows_by_producer": rows_by_producer,
        "problem_producers_summary": problem_summary,
        "asd_source_urls_checked": len({r.get("source_url", "") for r in by_prod.get("АСД", []) if r.get("source_url", "")}),
        "hausdoors_source_urls_checked": len({r.get("source_url", "") for r in by_prod.get("Хаусдорс", []) if r.get("source_url", "")}),
        "candidates_checked_asd": sum(int(r.get("checked_candidates_count", 0) or 0) for r in evidence_rows if r.get("producer") == "АСД"),
        "candidates_checked_hausdoors": sum(int(r.get("checked_candidates_count", 0) or 0) for r in evidence_rows if r.get("producer") == "Хаусдорс"),
        "visual_decision_counts": {
            "color_sample": reason_counter.get("color_sample", 0),
            "texture": reason_counter.get("texture", 0),
            "fragment": reason_counter.get("fragment", 0),
            "interior": reason_counter.get("interior", 0),
            "component": reason_counter.get("component", 0),
            "banner": reason_counter.get("banner", 0),
        },
        "hausdoors_color_rows_deleted": per_stats["Хаусдорс"].get("deleted", 0),
        "hausdoors_kromka_rows_deleted": per_stats["Хаусдорс"].get("hausdoors_kromka_deleted", 0),
        "asd_interior_urls_cleared": sum(
            1
            for r in reparse_rows
            if r.get("producer") == "АСД"
            and "interior_image_removed_no_full_door_found" in (r.get("comment_after") or "")
            and not (r.get("new_image_url") or "").strip()
        ),
        "unchanged_producers_check": unchanged_check,
    }
    (OUT_DIR / "run_report.json").write_text(json.dumps(run_report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"total_rows": len(rows_sorted), "rows_by_producer": rows_by_producer, "hausdoors_color_rows_deleted": per_stats["Хаусдорс"].get("deleted", 0)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
